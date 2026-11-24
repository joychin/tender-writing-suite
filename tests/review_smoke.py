#!/usr/bin/env python3
"""review_smoke.py — 评审台服务生命周期回归电池（跨平台，零第三方依赖）。

每条用例对应一个已坐实过的缺陷，走真实模块/命令行（同模型使用方式）。
背景：Windows 下 os.kill(pid, 0) 不是探活——信号 0 就是 CTRL_C_EVENT，走的是
GenerateConsoleCtrlEvent：目标不是进程组组长就抛 WinError 87（OSError），是组长
（宿主以新进程组起后台任务时）就等于给评审台发 Ctrl+C。2.2.0 真机日志里 --status
对活着的服务即因此 traceback。

R01 pid=0/负值/非整数不当作存活（Windows 下 os.kill(0,0) 会 WinError 87）
R02 os.kill 抛 OSError（Windows 真机形态）时 _pid_alive 返回 False，不抛
R03 已退出子进程的 pid 不当作存活
R04 _serve_url 拒绝缺字段/非法端口的残留 serve.json
R05 do_status 遇失效 serve.json：退出码 1、清理残留、不抛
R06 do_stop 遇 os.kill OSError：不抛、清理残留（收尾必做路径）
R07 真进程：起服务 → --status 识别运行中 → --stop 正常退出并清理
R08 Windows 分支确实走 _pid_alive_windows（OpenProcess），且查询异常不抛
R09 残留 serve.json 的 pid 被无关进程复用：--stop 先验身份（healthz 应答）再结束，
   不应答就只清记录——Windows 进程号回收快，收尾必做路径不得误杀别的程序
R10 环回探活不走代理：系统代理未豁免 127.0.0.1 时（Windows 注册表代理常见）
   start/--status/--stop 仍能探到服务，不误判「未运行」
"""
import io
import json
import os
import socket
import subprocess
import sys
import tempfile
import time
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
SERVE = ROOT / "tender-review" / "scripts" / "serve_review.py"
sys.path.insert(0, str(SERVE.parent))
import serve_review  # noqa: E402

PASS, FAIL = [], []


def check(name, cond, detail=""):
    if cond:
        PASS.append(name)
    else:
        FAIL.append(f"{name}{(' — ' + detail) if detail else ''}")


def run(name, fn):
    """一条用例出错只记它自己失败，后面的用例照跑。"""
    try:
        fn()
    except Exception as exc:  # noqa: BLE001
        FAIL.append(f"{name} — 用例异常 {type(exc).__name__}: {exc}")


def _serve_json(work: Path, pid: int = 123456, port: int = 17199) -> Path:
    d = work / "review"
    d.mkdir(parents=True, exist_ok=True)
    p = d / "serve.json"
    p.write_text(json.dumps({"work": str(work.resolve()), "port": port,
                             "token": "smoke", "pid": pid}), encoding="utf-8")
    return p


def _free_port() -> int:
    """先占后放：取一个此刻无人监听的环回端口。"""
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def t01():
    check("R01a pid=0 不存活", serve_review._pid_alive(0) is False)
    check("R01b 负 pid 不存活", serve_review._pid_alive(-1) is False)
    check("R01c 非整数 pid 不存活", serve_review._pid_alive("123") is False)
    check("R01d 超大 pid 不抛", serve_review._pid_alive(10 ** 20) is False)


def t02():
    with mock.patch.object(serve_review.os, "kill", side_effect=OSError(22, "参数错误")):
        got = serve_review._pid_alive(123456)
    check("R02 os.kill 抛 OSError 时按不存活处理", got is False)


def t03():
    p = subprocess.Popen([sys.executable, "-c", "pass"])
    p.wait()
    check("R03 已退出子进程 pid 不存活", serve_review._pid_alive(p.pid) is False)


def t04():
    check("R04a 合法 serve.json 出基址",
          serve_review._serve_url({"port": 17199, "token": "abc"})
          == "http://127.0.0.1:17199/abc/")
    check("R04b 缺 token 返回空", serve_review._serve_url({"port": 17199}) == "")
    check("R04c 端口 0 返回空", serve_review._serve_url({"port": 0, "token": "abc"}) == "")
    check("R04d 缺 port 返回空", serve_review._serve_url({"token": "abc"}) == "")


def t05():
    work = Path(tempfile.mkdtemp(prefix="review-smoke-status-"))
    spath = _serve_json(work)
    with mock.patch.object(serve_review.os, "kill", side_effect=OSError(22, "参数错误")), \
            redirect_stdout(io.StringIO()):
        rc = serve_review.do_status(work)
    check("R05a 失效 serve.json 退出码 1", rc == 1)
    check("R05b 残留 serve.json 被清理", not spath.exists())


def t06():
    work = Path(tempfile.mkdtemp(prefix="review-smoke-stop-"))
    spath = _serve_json(work)
    with mock.patch.object(serve_review, "_pid_alive", return_value=True), \
            mock.patch.object(serve_review, "probe", return_value=True), \
            mock.patch.object(serve_review.os, "kill", side_effect=OSError(22, "参数错误")), \
            mock.patch.object(serve_review.time, "sleep", return_value=None), \
            redirect_stdout(io.StringIO()):
        rc = serve_review.do_stop(work)
    check("R06a kill 抛 OSError 不崩", rc == 0)
    check("R06b 残留 serve.json 被清理", not spath.exists())


def t07():
    work = Path(tempfile.mkdtemp(prefix="review-smoke-live-"))
    env = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
    proc = subprocess.Popen(
        [sys.executable, "-B", str(SERVE), "--work", str(work), "--idle-timeout", "0"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
    spath = work / "review" / "serve.json"
    try:
        base = ""
        for _ in range(100):
            if spath.is_file():
                try:
                    cand = serve_review._serve_url(
                        json.loads(spath.read_text(encoding="utf-8")))
                    if cand and serve_review.probe(cand):
                        base = cand
                        break
                except Exception:  # noqa: BLE001
                    pass
            time.sleep(0.1)
        check("R07a 服务起来了", bool(base), "serve.json / healthz 未就绪")
        if base:
            r = subprocess.run(
                [sys.executable, "-B", str(SERVE), "--work", str(work), "--status"],
                capture_output=True, text=True, errors="replace", env=env)
            check("R07b --status 识别运行中",
                  r.returncode == 0 and "运行中" in r.stdout,
                  (r.stdout + r.stderr).strip())
            dead = {"http_proxy": "http://127.0.0.1:9", "HTTP_PROXY": "http://127.0.0.1:9",
                    "no_proxy": "", "NO_PROXY": ""}
            with mock.patch.dict(os.environ, dead), \
                    mock.patch.object(serve_review.urllib.request, "_opener", None):  # 丢掉已建的缺省 opener
                check("R10 代理指向无人监听端口时探活仍为真", serve_review.probe(base) is True)
        r = subprocess.run(
            [sys.executable, "-B", str(SERVE), "--work", str(work), "--stop"],
            capture_output=True, text=True, errors="replace", env=env)
        check("R07c --stop 正常退出", r.returncode == 0,
              (r.stdout + r.stderr).strip())
        check("R07d --stop 清理 serve.json", not spath.exists())
        for _ in range(30):
            if proc.poll() is not None:
                break
            time.sleep(0.1)
        check("R07e 服务进程已退出", proc.poll() is not None)
    finally:
        if proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(timeout=3)
            except subprocess.TimeoutExpired:
                proc.kill()


def t08():
    with mock.patch.object(serve_review.os, "name", "nt"), \
            mock.patch.object(serve_review, "_pid_alive_windows", return_value=True) as m:
        got = serve_review._pid_alive(123)
    check("R08a Windows 分支走 OpenProcess 查询", got is True)
    check("R08b 查询函数收到原 pid", m.call_args[0] == (123,))
    with mock.patch.object(serve_review.os, "name", "nt"), \
            mock.patch.object(serve_review, "_pid_alive_windows",
                              side_effect=OSError(22, "参数错误")):
        got = serve_review._pid_alive(123)
        check("R08c Windows 查询异常按不存活处理", got is False)


def t09():
    work = Path(tempfile.mkdtemp(prefix="review-smoke-reuse-"))
    bystander = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        spath = _serve_json(work, pid=bystander.pid, port=_free_port())
        with redirect_stdout(io.StringIO()):
            rc = serve_review.do_stop(work)
        time.sleep(0.3)
        check("R09a 被复用 pid 的无关进程不被 --stop 结束", bystander.poll() is None)
        check("R09b 残留 serve.json 被清理", not spath.exists())
        check("R09c 退出码 0", rc == 0)
    finally:
        if bystander.poll() is None:
            bystander.kill()
        bystander.wait()


for _name, _fn in [("R01", t01), ("R02", t02), ("R03", t03), ("R04", t04),
                   ("R05", t05), ("R06", t06), ("R07", t07), ("R08", t08), ("R09", t09)]:
    run(_name, _fn)

print(f"\n{len(PASS)} 通过 / {len(FAIL)} 失败")
for f in FAIL:
    print("  FAIL:", f)
sys.exit(1 if FAIL else 0)
