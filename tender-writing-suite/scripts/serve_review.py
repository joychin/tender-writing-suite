#!/usr/bin/env python3
"""serve_review.py — 评审台本地服务（环回 + token 路径，页面保存直写 work/ 产物）。

路由（全部挂在 /<token>/ 前缀下，无 token 一律 404）：
- GET  /<token>/          评审台页面（每次请求实时注入最新产物与草稿状态）
- GET  /<token>/healthz   心跳探活（页面 5s 一次；agent --status 探活）
- POST /<token>/apply     应用修改包（apply_review.apply_patch，含基底校验与备份）
- POST /<token>/draft     草稿暂存（防页面丢失编辑；{} = 清除）
- GET  /<token>/context   解析原文上下文切片（限 work/parse/ 内文件，防路径穿越）

稳定性设计：
- 端口占用：bind("127.0.0.1", 0) 由 OS 分配空闲端口，serve.json 记 {port,token,pid}；
  启动时探测存活+healthz+token 匹配则复用同一 URL，否则换端口重启；
- 进程被杀：服务无状态（GET 实时读盘、POST 立即落盘），重启零损失；页面 5s 心跳
  断开自动降级离线模式（修改保留本页，可导出修改包）；agent 用前 --status 探活；
- 僵尸进程：默认空闲 2 小时自动退出（--idle-timeout 0 关闭）；SIGTERM 清理 serve.json；
- 多会话：同一 work 目录共用一个服务（复用 serve.json），不同 work 各自独立。

安全声明：仅环回监听（127.0.0.1）、无外联；POST 校验 Origin/Referer 必须是本机；
context 只读 work/parse/ 下文件。 SkillHub 审查关注项均满足：无远程脚本、无网络外传。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import secrets
import signal
import socket
import sys
import threading
import time
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import build_review
from apply_review import apply_patch

SERVE_FILE = "review/serve.json"
DRAFT_FILE = "review/draft.json"
MAX_BODY = 2 * 1024 * 1024


def _read_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _pid_alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
        return True
    except (ProcessLookupError, PermissionError):
        return False


def probe(base: str, timeout: float = 1.5) -> bool:
    try:
        with urllib.request.urlopen(base + "healthz", timeout=timeout) as r:
            return r.status == 200
    except Exception:
        return False


# ---------- 服务 ----------

def make_handler(work: Path, token: str, state: dict):
    prefix = f"/{token}/"
    parse_dir = (work / "parse").resolve()

    class Handler(BaseHTTPRequestHandler):
        server_version = "ReviewWorkbench/1.3"

        def log_message(self, fmt, *args):  # 静默访问日志
            pass

        def _json(self, obj, code: int = 200) -> None:
            body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _reject(self, code: int, msg: str) -> None:
            self._json({"ok": False, "error": msg}, code)

        def _same_origin(self) -> bool:
            o = self.headers.get("Origin") or self.headers.get("Referer") or ""
            if not o:
                return True  # 无 Origin 头（非浏览器客户端）放行，token 仍是主防线
            m = re.match(r"^https?://(127\.0\.0\.1|localhost|\[::1\])(?::(\d+))?(?:/|$)", o)
            return bool(m and (m.group(2) is None or int(m.group(2)) == self.server.server_address[1]))

        def _guard(self, is_post: bool) -> bool:
            state["last"] = time.time()
            if not self.path.startswith(prefix):
                self._reject(404, "not found")
                return False
            if is_post:
                if not self._same_origin():
                    self._reject(403, "cross-origin request rejected")
                    return False
                if int(self.headers.get("Content-Length") or 0) > MAX_BODY:
                    self._reject(413, "payload too large")
                    return False
            return True

        def _body(self) -> bytes:
            return self.rfile.read(int(self.headers.get("Content-Length") or 0))

        def do_GET(self):
            if not self._guard(False):
                return
            path = self.path[len(prefix):].split("?")[0].lstrip("/")
            qs = dict(re.findall(r"([^?&=]+)=([^&]*)", self.path.split("?", 1)[-1])) if "?" in self.path else {}
            if path in ("", "index.html"):
                payload = build_review.collect_payload(work)
                draft = _read_json(work / DRAFT_FILE)
                payload["draft"] = draft if draft else None
                html = build_review.render_html(payload, api_base=prefix).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(html)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(html)
            elif path == "healthz":
                self._json({"ok": True, "work": str(work), "started": state["started"]})
            elif path == "context":
                q = {k: urllib.parse.unquote(v) for k, v in qs.items()}
                fname, line = q.get("file", ""), int(q.get("line", 0) or 0)
                target = (parse_dir / fname).resolve() if fname else None
                if not fname or not target or parse_dir not in target.parents or not target.is_file():
                    self._json({"ok": False, "error": "文件不在解析产物范围内或不存在"}, 404)
                    return
                lines = target.read_text(encoding="utf-8", errors="replace").split("\n")
                lo, hi = max(1, line - 8), min(len(lines), line + 12)
                self._json({"ok": True, "file": fname, "from": lo, "to": hi,
                            "lines": [{"no": i, "text": lines[i - 1]} for i in range(lo, hi + 1)]})
            else:
                self._reject(404, "not found")

        def do_POST(self):
            if not self._guard(True):
                return
            path = self.path[len(prefix):].lstrip("/")
            if path == "apply":
                try:
                    patch = json.loads(self._body().decode("utf-8") or "{}")
                except Exception as e:
                    self._json({"ok": False, "errors": [f"修改包不是合法 JSON：{e}"]}, 400)
                    return
                result = apply_patch(work, patch)
                self._json(result, 200 if result.get("ok") else 409 if result.get("stage") == "base" else 400)
            elif path == "draft":
                try:
                    body = json.loads(self._body().decode("utf-8") or "{}")
                except Exception:
                    body = {}
                if body:
                    body["at"] = time.strftime("%H:%M")
                    _write_json(work / DRAFT_FILE, body)
                else:
                    (work / DRAFT_FILE).unlink(missing_ok=True)
                self._json({"ok": True})
            else:
                self._reject(404, "not found")

    return Handler


def start(work: Path, idle_timeout: int) -> tuple[str, bool]:
    """启动（或复用）服务；返回 URL 与是否新起。"""
    import urllib.parse
    work = work.resolve()
    spath = work / SERVE_FILE
    old = _read_json(spath) if spath.is_file() else None
    if old:
        base = f"http://127.0.0.1:{old['port']}/{old['token']}/"
        if _pid_alive(old.get("pid", 0)) and probe(base):
            return base, False, None
    token = secrets.token_hex(8)
    state = {"last": time.time(), "started": time.strftime("%Y-%m-%d %H:%M")}
    srv = ThreadingHTTPServer(("127.0.0.1", 0), make_handler(work, token, state))
    port = srv.server_address[1]
    url = f"http://127.0.0.1:{port}/{token}/"
    _write_json(spath, {"port": port, "token": token, "pid": os.getpid(),
                        "started": state["started"], "work": str(work)})
    t = threading.Thread(target=idle_watchdog, args=(srv, state, idle_timeout, spath), daemon=True)
    t.start()

    def bye(signum, frame):
        spath.unlink(missing_ok=True)
        os._exit(0)
    signal.signal(signal.SIGTERM, bye)
    signal.signal(signal.SIGINT, bye)
    serve_t = threading.Thread(target=srv.serve_forever, kwargs={"poll_interval": 0.5}, daemon=True)
    serve_t.start()
    return url, True, serve_t


def idle_watchdog(srv: ThreadingHTTPServer, state: dict, idle_timeout: int, spath: Path) -> None:
    if idle_timeout <= 0:
        return
    while True:
        time.sleep(30)
        if time.time() - state["last"] > idle_timeout:
            spath.unlink(missing_ok=True)
            srv.shutdown()
            break


def do_status(work: Path) -> int:
    old = _read_json(work / SERVE_FILE) if (work / SERVE_FILE).is_file() else None
    if old and _pid_alive(old.get("pid", 0)) and probe(f"http://127.0.0.1:{old['port']}/{old['token']}/"):
        print(f"[评审台服务运行中] http://127.0.0.1:{old['port']}/{old['token']}/")
        return 0
    (work / SERVE_FILE).unlink(missing_ok=True)
    print("[评审台服务未运行]")
    return 1


def do_stop(work: Path) -> int:
    old = _read_json(work / SERVE_FILE) if (work / SERVE_FILE).is_file() else None
    if not old:
        print("[评审台服务未运行]")
        return 0
    pid = old.get("pid", 0)
    if _pid_alive(pid):
        try:
            os.kill(pid, signal.SIGTERM)
            for _ in range(20):
                if not _pid_alive(pid):
                    break
                time.sleep(0.1)
            else:
                os.kill(pid, signal.SIGKILL)
        except (ProcessLookupError, PermissionError):
            pass
        print(f"[评审台服务已停止] pid={pid}")
    else:
        print("[评审台服务未运行]（清理残留记录）")
    (work / SERVE_FILE).unlink(missing_ok=True)
    return 0


def main() -> None:
    import urllib.parse  # noqa: F401  (handler 里 unquote 用)
    global urllib
    ap = argparse.ArgumentParser(description="评审台本地服务（环回直写 work/ 产物）")
    ap.add_argument("--work", default="work", help="工作目录（默认 work）")
    ap.add_argument("--idle-timeout", type=int, default=7200, help="空闲自动退出秒数（默认 7200，0=不超时）")
    ap.add_argument("--status", action="store_true", help="探活：打印运行中的服务 URL")
    ap.add_argument("--stop", action="store_true", help="停止运行中的服务（收尾调用）")
    args = ap.parse_args()
    work = Path(args.work)
    if args.status:
        sys.exit(do_status(work))
    if args.stop:
        sys.exit(do_stop(work))
    url, fresh, serve_t = start(work, args.idle_timeout)
    print(f"[评审台服务{'已启动' if fresh else '复用中'}] {url}", flush=True)
    print(f"工作目录：{work.resolve()}", flush=True)
    print("空闲超时自动退出（--idle-timeout 0 关闭）；收尾时用 --stop 停止。", flush=True)
    if not fresh:
        return  # 复用模式：打印 URL 即退出，服务由原进程持有
    try:
        serve_t.join()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
