# LibreOffice 상주 변환기 — webviewer(src/converters/loWorker.ts)가 LibreOffice에 딸린 python.exe로 띄운다.
#
# 매 변환마다 soffice.exe를 새로 띄우면 시작에만 1.5초 남짓이 든다. 이 스크립트는 soffice를 한 번 띄워
# UNO로 붙어 있다가, 표준 입력으로 한 줄씩 들어오는 변환 요청(JSON)을 처리하고 결과를 한 줄로 돌려준다.
#   요청: {"id": 1, "src": "C:\\...\\a.docx", "out": "C:\\...\\a.pdf", "filterData": {"SinglePageSheets": true}}
#   응답: {"id": 1, "ok": true, "ms": 312}  또는  {"id": 1, "ok": false, "error": "..."}
#   변환마다 뒷정리가 끝나면 {"idle": true} — webviewer는 그때까지 새 변환을 맡기지 않는다(예전 방식으로 돌린다)
# 표준 입력이 닫히면(webviewer 종료) soffice도 끄고 끝낸다.
#
# 발표(pptx) 문서는 닫지 않는다: 이 LibreOffice(Windows, 원격 연결)에서 발표 문서를 닫으면 close·dispose 어느 쪽이든
# soffice가 9초쯤 멈췄다가 죽었다(실측, 2026-10-19). 대신 열어 둔 채 쌓아 두고, 일정 개수·크기가 되면 한가할 때
# soffice를 새로 띄운다. 원본 파일을 붙잡지 않게 발표 문서는 파일이 아니라 메모리(스트림)로 연다.
import json
import os
import subprocess
import sys
import time
import traceback

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException

# 표준 입출력은 UTF-8로 — 한국어 Windows의 기본(cp949)으로 읽으면 한글 파일 이름이 깨져 "type detection failed"로
# 실패했다(실측: "한글 보고서(최종).docx"). webviewer는 UTF-8로 보내고 UTF-8로 읽는다.
sys.stdin.reconfigure(encoding="utf-8")
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

SOFFICE, PIPE_BASE, PROFILE_URL = sys.argv[1], sys.argv[2], sys.argv[3]
# 띄울 때마다 새 파이프 이름 — 같은 이름이면 이 프로세스의 UNO가 끊긴 예전 연결을 다시 꺼내 써서, 다시 띄운 soffice에
# 바로 "bridge disposed"로 실패했다(실측: 새로 띄우기 3번 모두 즉시 실패)
pipe_name = PIPE_BASE
start_count = 0
PROFILE_DIR = uno.fileUrlToSystemPath(PROFILE_URL)
# 지난번 변환기가 강제로 꺼져 남은 soffice — 같은 프로필을 붙잡고 있으면 새로 띄운 soffice가 그쪽으로 넘어가 버린다
PIDFILE = os.path.join(PROFILE_DIR, "worker.pid")

IMPRESS_EXTS = {"pptx", "ppt", "pps", "ppsx", "odp"}
# 열어 둔 발표 문서가 이만큼 쌓이면 한가할 때 soffice를 새로 띄운다
MAX_OPEN_IMPRESS = 8
MAX_OPEN_IMPRESS_BYTES = 30 * 1024 * 1024

proc = None
ctx = None
desktop = None
open_impress = []  # 닫지 않고 둔 발표 문서
open_impress_bytes = 0
T0 = time.time()


def log(message):
    # 진단용 — 표준 오류로(webviewer는 버린다). 직접 실행해 볼 때 단계별 시간을 본다
    sys.stderr.write("[%6.1fs] %s\n" % (time.time() - T0, message))
    sys.stderr.flush()


def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def kill_tree(pid):
    # soffice.exe는 soffice.bin을 따로 띄운다 — 트리째 끄지 않으면 soffice.bin이 남는다
    if os.name == "nt":
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        try:
            os.kill(pid, 9)
        except OSError:
            pass


def kill_profile_processes():
    """이 프로필을 쓰는 soffice를 전부 끈다 — 처음 띄운 soffice.exe의 자식만 끄면, 스스로 다시 켜진 soffice.bin
    (새 프로필을 만들 때 한 번 그렇게 된다)은 그 트리 밖이라 남았다. 명령줄의 프로필 폴더 이름으로 찾는다."""
    if os.name != "nt":
        return
    name = os.path.basename(PROFILE_DIR.rstrip("\\/"))
    # 폴더 이름 바로 뒤가 영숫자가 아닐 때만 — worker-8000이 worker-80001까지 잡지 않게
    script = (
        "Get-CimInstance Win32_Process -Filter \"Name like 'soffice%'\" | "
        "Where-Object { $_.CommandLine -match ([regex]::Escape('" + name + "') + '(?![0-9A-Za-z])') } | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    )
    subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def kill_leftover():
    # 지난번 변환기가 강제로 꺼져 남은 soffice가 이 프로필을 붙잡고 있으면 새 soffice가 그쪽으로 넘어가 버린다
    kill_profile_processes()


def stop():
    global proc, desktop, ctx, open_impress, open_impress_bytes
    # 열어 둔 발표 문서가 없으면 점잖게 끈다. 있으면 terminate가 그걸 닫다가 멈추므로(위 머리말) 바로 끈다.
    if desktop is not None and proc is not None and proc.poll() is None and not open_impress:
        try:
            desktop.terminate()
            proc.wait(timeout=5)
        except Exception:
            pass
    desktop = None
    ctx = None
    open_impress = []
    open_impress_bytes = 0
    if proc is not None:
        if proc.poll() is None:
            kill_tree(proc.pid)
        proc = None
    kill_profile_processes()
    try:
        os.remove(PIDFILE)
    except OSError:
        pass


def connect():
    global desktop, ctx
    desktop = None
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(360):  # 최대 90초 — 새 프로필을 만들 때는 30초 넘게 걸렸다(실측 37초)
        try:
            ctx = resolver.resolve("uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipe_name)
            desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            return
        except NoConnectException:
            # soffice.exe가 끝나도 기다린다 — 새 프로필이면 soffice가 한 번 스스로 다시 켜지는데, 그때 soffice.exe는
            # 끝나고 새 soffice.bin이 같은 파이프로 다시 받는다(그때 포기해서 새 프로필로는 변환기가 켜지지 않았다)
            time.sleep(0.25)
    raise RuntimeError("LibreOffice에 연결하지 못했습니다")


def start():
    # 문서를 열어 둔 채 강제로 끈 직후에 띄우면, soffice가 프로필을 되살리며 웜업 도중 한 번 끊기곤 했다(실측) —
    # 그러면 한 번 더 깨끗하게 띄운다
    for attempt in range(3):
        try:
            start_once()
            return
        except Exception as error:
            log("띄우기 %d번째 실패: %s" % (attempt + 1, error))
            if attempt == 2:
                raise
            time.sleep(2)


def start_once():
    global proc, pipe_name, start_count
    start_count += 1
    pipe_name = "%s_%d" % (PIPE_BASE, start_count)
    kill_leftover()
    stop()
    os.makedirs(PROFILE_DIR, exist_ok=True)
    proc = subprocess.Popen(
        [SOFFICE, "--headless", "--norestore", "-env:UserInstallation=" + PROFILE_URL, "--accept=pipe,name=%s;urp;StarOffice.ComponentContext" % pipe_name]
    )
    with open(PIDFILE, "w") as f:
        f.write(str(proc.pid))
    log("soffice 띄움 pid=%d" % proc.pid)
    connect()
    log("연결됨")
    # 글·발표·표를 한 번씩 실제 파일로 바꿔 둔다 — 첫 사용자 변환이 각 형식의 첫 적재 비용을 치르지 않게
    warm_pdf = os.path.join(PROFILE_DIR, "warmup.pdf")
    for factory, save_filter, name in (
        ("swriter", "MS Word 2007 XML", "warmup.docx"),
        ("simpress", "Impress MS PowerPoint 2007 XML", "warmup.pptx"),
        ("scalc", "Calc MS Excel 2007 XML", "warmup.xlsx"),
    ):
        path = os.path.join(PROFILE_DIR, name)
        blank = desktop.loadComponentFromURL("private:factory/" + factory, "_blank", 0, (pv("Hidden", True),))
        uno.invoke(blank, "storeToURL", (uno.systemPathToFileUrl(path), (pv("FilterName", save_filter),)))
        if factory == "simpress":
            keep_open(blank, 0)
        else:
            close_quietly(blank)
        finish(convert(path, warm_pdf, None), path)
    desktop.getComponents()
    log("웜업 끝")


def export_filter(doc):
    if doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        return "calc_pdf_Export"
    if doc.supportsService("com.sun.star.presentation.PresentationDocument"):
        return "impress_pdf_Export"
    if doc.supportsService("com.sun.star.drawing.DrawingDocument"):
        return "draw_pdf_Export"
    return "writer_pdf_Export"


def close_quietly(doc):
    try:
        doc.close(True)
    except Exception:
        pass


def is_impress(path):
    return os.path.splitext(path)[1].lower().lstrip(".") in IMPRESS_EXTS


def keep_open(doc, size):
    global open_impress_bytes
    open_impress.append(doc)
    open_impress_bytes += size


def finish(doc, src):
    """변환이 끝난 문서 — 발표 문서는 닫지 않고 둔다(머리말). 나머지는 닫는다."""
    if is_impress(src):
        keep_open(doc, os.path.getsize(src) if os.path.exists(src) else 0)
    else:
        close_quietly(doc)


def load(src):
    if is_impress(src):
        # 메모리로 연다 — 열어 둔 채 쌓아 두므로, 파일로 열면 원본 사본(사용자 문서)이 잠겨 지울 수 없다
        with open(src, "rb") as f:
            data = f.read()
        stream = ctx.ServiceManager.createInstanceWithArgumentsAndContext(
            "com.sun.star.io.SequenceInputStream", (uno.ByteSequence(data),), ctx
        )
        return desktop.loadComponentFromURL("private:stream", "_blank", 0, (pv("Hidden", True), pv("InputStream", stream)))
    return desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))


def convert(src, out, filter_data):
    """변환하고 열린 문서를 돌려준다 — 닫기(또는 열어 두기)는 호출부가 응답을 보낸 뒤에 한다."""
    doc = load(src)
    if doc is None:
        if os.environ.get("DOCVIEW_LO_WORKER_DEBUG") == "1":
            # 진단용 사본(디버그 때만) — 운영에선 사용자 문서를 남기지 않는다
            import shutil
            shutil.copy(src, os.path.join(PROFILE_DIR, "failed-" + os.path.basename(src)))
        raise RuntimeError("문서를 열지 못했습니다")
    try:
        props = [pv("FilterName", export_filter(doc))]
        if filter_data:
            data = uno.Any("[]com.sun.star.beans.PropertyValue", tuple(pv(k, v) for k, v in filter_data.items()))
            props.append(pv("FilterData", data))
        uno.invoke(doc, "storeToURL", (uno.systemPathToFileUrl(out), tuple(props)))
    except Exception:
        finish(doc, src)
        raise
    return doc


def healthy():
    try:
        desktop.getComponents()
        return proc is not None and proc.poll() is None
    except Exception:
        return False


def reply(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def main():
    try:
        start()
    except Exception as error:
        log("시작 실패: %s" % error)
        reply({"ready": False, "error": str(error)})
        stop()
        return
    reply({"ready": True})
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        job = json.loads(line)
        began = time.time()
        doc = None
        try:
            doc = convert(job["src"], job["out"], job.get("filterData"))
            reply({"id": job["id"], "ok": True, "ms": int((time.time() - began) * 1000)})
        except Exception as error:
            reply({"id": job["id"], "ok": False, "error": "%s: %s" % (type(error).__name__, error), "trace": traceback.format_exc()[-800:]})
        # 뒷정리 — 응답을 보낸 뒤라 사용자는 기다리지 않는다. webviewer는 "idle"까지 새 변환을 맡기지 않는다.
        if doc is not None:
            finish(doc, job["src"])
        restarted = None
        if not healthy():
            restarted = "soffice가 응답하지 않아 새로 띄움"
        elif len(open_impress) >= MAX_OPEN_IMPRESS or open_impress_bytes >= MAX_OPEN_IMPRESS_BYTES:
            restarted = "열어 둔 발표 문서 %d개(%dMB) — 메모리를 돌려받으려고 새로 띄움" % (len(open_impress), open_impress_bytes // (1024 * 1024))
        if restarted:
            log(restarted)
            try:
                start()
            except Exception as error:
                reply({"idle": True, "restarted": restarted, "fatal": str(error)})
                stop()
                return
        reply({"idle": True, "restarted": restarted})
    stop()


if __name__ == "__main__":
    main()
