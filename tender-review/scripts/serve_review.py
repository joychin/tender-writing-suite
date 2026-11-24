#!/usr/bin/env python3
"""serve_review.py — 评审台本地服务（环回 + token 路径，页面保存直写 work/ 产物）。

路由（全部挂在 /<token>/ 前缀下，无 token 一律 404）：
- GET  /<token>/          评审台页面（每次请求实时注入最新产物、草稿状态与产物指纹；
                           评审台.html 为同一页面的别名——分析报告页脚回链走相对路径用）
- GET  /<token>/report.html  分析报告页（实时渲染；报告.html 为别名；评审台总览的「打开报告页面」入口指向这里）
- GET  /<token>/progress.html  写作进度面板（实时渲染+指纹门控自刷新；进度.html 为别名；写作阶段监视面）
- GET  /<token>/healthz   心跳探活+产物指纹（评审台 5s 一次比对出「材料已更新」提示；
                           进度面板 30s 一次驱动门控自刷新；agent --status 探活）
- POST /<token>/apply     应用修改包（apply_review.apply_patch，含基底校验与备份；
                           回包带应用后的新指纹，页面基线同步、不把自己的保存误报成外部更新）
- POST /<token>/draft     草稿暂存（防页面丢失编辑；{} = 清除）
- GET  /<token>/context   解析原文上下文切片（限 work/parse/ 内文件，防路径穿越）

稳定性设计：
- 端口占用：bind("127.0.0.1", 0) 由 OS 分配空闲端口，serve.json 记 {port,token,pid}；
  启动时探测存活+healthz+token 匹配则复用同一 URL，否则换端口重启；
- 进程被杀：服务无状态（GET 实时读盘、POST 立即落盘），重启零损失；页面 5s 心跳
  断开自动降级离线模式（修改保留本页，可导出修改包）；agent 用前 --status 探活；
- 僵尸进程：默认空闲 2 小时自动退出（--idle-timeout 0 关闭）；SIGTERM 清理 serve.json
  （Windows 没有 SIGTERM 投递，--stop 是强制结束，记录由 --stop 自己清）；
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


def _serve_url(old: dict) -> str:
    """serve.json → 服务基址；字段缺失或非法返回空串（旧版/损坏的残留文件）。"""
    port, token = old.get("port"), old.get("token")
    if isinstance(port, int) and 0 < port < 65536 and isinstance(token, str) and token:
        return f"http://127.0.0.1:{port}/{token}/"
    return ""


def _pid_alive_windows(pid: int) -> bool:
    """Windows 探活：OpenProcess 只读查询。

    Windows 下 os.kill(pid, 0) 不是探活：信号 0 就是 CTRL_C_EVENT，走的是
    GenerateConsoleCtrlEvent——目标不是进程组组长就抛 WinError 87（OSError），
    是组长（宿主以新进程组起后台任务时）就等于给评审台发 Ctrl+C 把它关掉。
    必须走只读的进程查询。
    """
    import ctypes
    from ctypes import wintypes

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    STILL_ACTIVE = 259
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not handle:
        return False
    try:
        code = wintypes.DWORD()
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
            return False
        return code.value == STILL_ACTIVE
    finally:
        kernel32.CloseHandle(handle)


def _pid_alive(pid: int) -> bool:
    """pid 有效且进程存活；探测失败一律按「不存活」处理，绝不抛给调用方。

    Windows 用 OpenProcess 查询（os.kill(pid, 0) 在 Windows 不是探活）；
    Unix 用 os.kill(pid, 0)。死 pid / 权限不足 / 非法 pid 都返回 False。
    """
    if not isinstance(pid, int) or pid <= 0:
        return False
    try:
        if os.name == "nt":
            return _pid_alive_windows(pid)
        os.kill(pid, 0)
        return True
    except Exception:  # noqa: BLE001 — 探活失败一律按不存活，绝不让 --status/--stop 崩
        return False


# 环回探活专用 opener：显式不带代理。urlopen 缺省会读系统代理（Windows 读注册表、
# macOS 读系统设置），代理豁免名单没含 127.0.0.1 时探活会被转给代理——本地服务
# 明明活着却判「未运行」（--status 删 serve.json、start 重复起服务）。
_PROBE = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def probe(base: str, timeout: float = 1.5) -> bool:
    try:
        with _PROBE.open(base + "healthz", timeout=timeout) as r:
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

        def _html(self, html: str) -> None:
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

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
                try:
                    clen = int(self.headers.get("Content-Length") or 0)
                except ValueError:
                    clen = 0
                if clen > MAX_BODY:
                    self._reject(413, "payload too large")
                    return False
            return True

        def _body(self) -> bytes:
            return self.rfile.read(int(self.headers.get("Content-Length") or 0))

        def do_GET(self):
            # 浏览器向源站根路径要 /favicon.ico（不带 token 前缀）——在守卫前静默
            # 应答 204，否则 console 会记一条 404 噪声
            if urllib.parse.urlsplit(self.path).path.endswith("/favicon.ico"):
                self.send_response(204)
                self.end_headers()
                return
            if not self._guard(False):
                return
            # 报告页脚回链用的是相对路径「评审台.html」，URL 里以百分号编码到达，先解码再比对
            path = urllib.parse.unquote(self.path[len(prefix):].split("?")[0].lstrip("/"))
            qs = {k: v[0] for k, v in urllib.parse.parse_qs(self.path.split("?", 1)[-1]).items()} if "?" in self.path else {}
            if path in ("", "index.html", "评审台.html"):
                try:
                    payload = build_review.collect_payload(work)
                    draft = _read_json(work / DRAFT_FILE)
                    payload["draft"] = draft if draft else None
                    payload["fp"] = build_review.fingerprint(work)
                    html = build_review.render_html(payload, api_base=prefix)
                except Exception as exc:
                    self._json({"ok": False, "error": f"产物读取失败：{exc}"}, 500)
                    return
                self._html(html)
            elif path in ("report.html", "招标文件分析报告.html"):
                try:
                    html = build_review.render_report(build_review.collect_payload(work), workbench=True)
                except Exception as exc:
                    self._json({"ok": False, "error": f"产物读取失败：{exc}"}, 500)
                    return
                self._html(html)
            elif path in ("progress.html", "写作进度.html"):
                try:
                    pp = build_review.progress_payload(work)
                    pp["fp"] = build_review.fingerprint(work)
                    html = build_review.render_progress(pp, live=True)
                except Exception as exc:
                    self._json({"ok": False, "error": f"产物读取失败：{exc}"}, 500)
                    return
                self._html(html)
            elif path == "healthz":
                self._json({"ok": True, "work": str(work), "started": state["started"],
                            "fp": build_review.fingerprint(work)})
            elif path == "context":
                fname = qs.get("file", "")
                try:
                    line = int(qs.get("line", 0) or 0)
                except ValueError:
                    line = 0
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
                try:
                    result = apply_patch(work, patch)
                except Exception as exc:
                    self._json({"ok": False, "stage": "crash",
                                "errors": [f"应用异常（产物保持原状或已有备份）：{exc}"]}, 500)
                    return
                if result.get("ok"):
                    (work / DRAFT_FILE).unlink(missing_ok=True)  # 应用成功，待存草稿已过时
                    result["fp"] = build_review.fingerprint(work)  # 新基线：防页面把自己的保存误报成外部更新
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


def start(work: Path, idle_timeout: int) -> tuple[str, bool, object]:
    """启动（或复用）服务；返回 (URL, 是否新起, 服务线程)。
    复用仅限同一工作目录（serve.json 记 work 比对）——复制工作区会把旧
    serve.json 一起带走，误复用会让两个目录指向同一个服务（实测踩坑）。"""
    work = work.resolve()
    spath = work / SERVE_FILE
    old = _read_json(spath) if spath.is_file() else None
    if old and old.get("work") == str(work):
        base = _serve_url(old)
        if base and _pid_alive(old.get("pid", 0)) and probe(base):
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
    work = work.resolve()
    old = _read_json(work / SERVE_FILE) if (work / SERVE_FILE).is_file() else None
    base = _serve_url(old) if old and old.get("work") == str(work) else ""
    if base and _pid_alive(old.get("pid", 0)) and probe(base):
        print(f"[评审台服务运行中] {base}")
        return 0
    (work / SERVE_FILE).unlink(missing_ok=True)
    print("[评审台服务未运行]")
    return 1


def do_stop(work: Path) -> int:
    work = work.resolve()
    old = _read_json(work / SERVE_FILE) if (work / SERVE_FILE).is_file() else None
    if not old or old.get("work") != str(work):
        print("[评审台服务未运行]")
        (work / SERVE_FILE).unlink(missing_ok=True)
        return 0
    pid = old.get("pid", 0)
    base = _serve_url(old)
    if _pid_alive(pid) and not (base and probe(base)):
        # 进程号还在，但不是评审台在应答：残留记录的 pid 被别的程序复用了（Windows
        # 进程号回收很快）或服务已卡死。结束进程不可撤销，身份没验实就不动它。
        print(f"[评审台服务未停止] 记录的进程 pid={pid} 仍在，但不回应评审台探活——"
              "多半是该进程号已被其他程序占用，未强制结束（防误杀）。残留记录已清理；"
              "评审台服务若仍在运行，会在空闲超时后自行退出")
    elif _pid_alive(pid):
        try:
            os.kill(pid, signal.SIGTERM)  # Windows 下是强制结束（无 SIGTERM 投递），记录由本命令清
        except OSError:
            pass
        for _ in range(20):
            if not _pid_alive(pid):
                break
            time.sleep(0.1)
        else:
            sigkill = getattr(signal, "SIGKILL", None)  # Windows 无 SIGKILL
            if sigkill is not None:
                try:
                    os.kill(pid, sigkill)
                except OSError:
                    pass
        if _pid_alive(pid):
            print(f"[评审台服务未停止] pid={pid}（权限不足或进程无响应，请手动结束）")
        else:
            print(f"[评审台服务已停止] pid={pid}")
    else:
        print("[评审台服务未运行]（清理残留记录）")
    (work / SERVE_FILE).unlink(missing_ok=True)
    return 0


def main() -> None:
    for _s in (sys.stdout, sys.stderr):
        if hasattr(_s, "reconfigure"):
            _s.reconfigure(errors="replace")  # 中文 Windows 管道（GBK）下 ⚠️ 会编码崩，降级输出
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
    print("重复交付同一页面时，链接须带版本参数 ?v=<产物指纹前8位>（宿主预览面板对相同"
          "URL 不重新加载，带参才强制刷新）——每道门重跑 build_review 后复制其回显的"
          "[交付链接]，不要复用起台时的裸 URL。", flush=True)
    print(f"写作进度面板：{url}progress.html（第五步监视面）——第五步首节开写前的第一动作是"
          "跑 build_review.py 加 --open-progress 首开面板并把链接发给用户，"
          "未首开写节会被 docx_toolkit / check_writing --section 硬拦。", flush=True)
    print("[停轮提醒] 起台≠门已过：把上面的 URL 以 Markdown 链接（[打开评审页面](URL)）"
          "发给用户并停轮等答复——链接之后本轮不得再发起推进动作"
          "（用户已明确授权连续模式除外）。"
          "用户回「继续/先这样」只越过本门等待、按默认方案推进到下一门产物落盘为止，"
          "下一门交付后照常停轮——门不得同轮连过；承诺卡未呈现（G4/G5 合并门）前，"
          "任何节不得落承诺数字（全部待补 + 批注）。门指针里带上「需要你补充」表格"
          "（| 需补充 | 影响 | 不补充的处理 |），并固定告知：不补充也可以继续往下——"
          "缺失部分会以 Word 批注标在正文对应位置。", flush=True)
    if not fresh:
        return  # 复用模式：打印 URL 即退出，服务由原进程持有
    try:
        serve_t.join()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
