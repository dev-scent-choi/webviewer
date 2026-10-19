# LibreOffice 상주 변환기 — webviewer(src/converters/loWorker.ts)가 LibreOffice에 딸린 python.exe로 띄운다.
#
# 매 변환마다 soffice.exe를 새로 띄우면 시작에만 1.5초 남짓이 든다. 이 스크립트는 soffice를 한 번 띄워
# UNO로 붙어 있다가, 표준 입력으로 한 줄씩 들어오는 변환 요청(JSON)을 처리하고 결과를 한 줄로 돌려준다.
#   요청: {"id": 1, "src": "C:\\...\\a.docx", "out": "C:\\...\\a.pdf", "filterData": {"SinglePageSheets": true}}
#   응답: {"id": 1, "ok": true, "ms": 312}  또는  {"id": 1, "ok": false, "error": "..."}
# 표준 입력이 닫히면(webviewer 종료) soffice도 끄고 끝낸다.
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

SOFFICE, PIPE_NAME, PROFILE_URL = sys.argv[1], sys.argv[2], sys.argv[3]
UNO_URL = "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % PIPE_NAME

proc = None
desktop = None
# 지난번 변환기가 강제로 꺼져 남은 soffice — 같은 프로필을 붙잡고 있으면 새로 띄운 soffice가 그쪽으로 넘어가 버린다
PIDFILE = os.path.join(uno.fileUrlToSystemPath(PROFILE_URL), "worker.pid")


def kill_tree(pid):
    # soffice.exe는 soffice.bin을 따로 띄운다 — 트리째 끄지 않으면 soffice.bin이 남는다
    if os.name == "nt":
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        try:
            os.kill(pid, 9)
        except OSError:
            pass


def kill_leftover():
    try:
        with open(PIDFILE) as f:
            pid = int(f.read().strip())
    except (OSError, ValueError):
        return
    if os.name == "nt":
        listed = subprocess.run(["tasklist", "/FI", "PID eq %d" % pid, "/NH"], capture_output=True, text=True).stdout.lower()
        if "soffice" not in listed:
            return  # 이미 없거나 다른 프로그램이 그 번호를 쓰고 있다
    kill_tree(pid)


def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def stop():
    global proc, desktop
    # 먼저 점잖게 끈다 — 강제로 끄면 다음에 켤 때 soffice가 프로필을 되살리느라 첫 변환 뒤 스스로 한 번 다시
    # 켜졌다(그 변환이 10초 남짓 걸림, 실측). 5초 안에 안 꺼지면 그때 강제로.
    if desktop is not None and proc is not None and proc.poll() is None:
        try:
            desktop.terminate()
            proc.wait(timeout=5)
        except Exception:
            pass
    desktop = None
    if proc is not None:
        if proc.poll() is None:
            kill_tree(proc.pid)
        proc = None
    try:
        os.remove(PIDFILE)
    except OSError:
        pass


def start():
    global proc, desktop
    kill_leftover()
    stop()
    os.makedirs(os.path.dirname(PIDFILE), exist_ok=True)
    proc = subprocess.Popen(
        [
            SOFFICE,
            "--headless",
            "--invisible",
            "--norestore",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "-env:UserInstallation=" + PROFILE_URL,
            "--accept=pipe,name=%s;urp;StarOffice.ComponentContext" % PIPE_NAME,
        ]
    )
    with open(PIDFILE, "w") as f:
        f.write(str(proc.pid))
    connect()
    # 처음 한 번 빈 문서를 열고 닫아 본다 — 새(또는 막 고쳐진) 프로필이면 soffice가 이 무렵 스스로 한 번 다시
    # 켜지는데(soffice.exe는 그대로, 안쪽 soffice.bin만 새로 — "URP bridge already disposed"), 그걸 첫 변환이 아니라
    # 여기서 겪고 다시 붙는다.
    # 실제로 PDF까지 내보내야 그 다시 켜짐이 일어난다(빈 문서를 열고 닫기만 해서는 첫 실제 변환 때 일어났다 — 실측)
    warm_pdf = os.path.join(os.path.dirname(PIDFILE), "warmup.pdf")
    # 글·발표·표 세 종류를 모두 한 번씩 — 처음 쓰는 종류의 변환 뒤에 다시 켜지는 일이 있었다(실측: 발표 다음 첫 글 문서)
    for _ in range(3):
        try:
            for factory, export in (("swriter", "writer_pdf_Export"), ("simpress", "impress_pdf_Export"), ("scalc", "calc_pdf_Export")):
                doc = desktop.loadComponentFromURL("private:factory/" + factory, "_blank", 0, (pv("Hidden", True),))
                uno.invoke(doc, "storeToURL", (uno.systemPathToFileUrl(warm_pdf), (pv("FilterName", export),)))
                close_quietly(doc)
            time.sleep(1.5)  # 다시 켜짐은 내보낸 직후 조금 뒤에 일어난다
            desktop.getComponents()  # 연결이 살아 있는지
            return
        except Exception:
            time.sleep(1)
            connect()


def connect():
    """soffice.exe는 살아 있는데 연결만 끊긴 경우(스스로 다시 켜짐)에도 쓴다 — 죽이고 새로 띄우면 다시 켜지는 중인
    프로필을 또 건드려 10초씩 걸리고 같은 일이 되풀이됐다(실측)."""
    global desktop
    desktop = None
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(240):  # 최대 60초 — 첫 프로필 생성은 10초 넘게 걸리기도 한다
        try:
            ctx = resolver.resolve(UNO_URL)
            desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            return
        except NoConnectException:
            if proc is None or proc.poll() is not None:
                break
            time.sleep(0.25)
    raise RuntimeError("LibreOffice에 연결하지 못했습니다")


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


def convert(src, out, filter_data):
    """변환하고 열린 문서를 돌려준다 — 닫기는 호출부가 응답을 보낸 뒤에 한다. 그림이 많은 큰 pptx는
    닫는 데만 변환만큼(실측 8초 남짓) 걸려, 닫은 뒤에 응답하면 그만큼 사용자가 더 기다렸다."""
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))
    if doc is None:
        raise RuntimeError("문서를 열지 못했습니다")
    try:
        props = [pv("FilterName", export_filter(doc))]
        if filter_data:
            data = uno.Any("[]com.sun.star.beans.PropertyValue", tuple(pv(k, v) for k, v in filter_data.items()))
            props.append(pv("FilterData", data))
        uno.invoke(doc, "storeToURL", (uno.systemPathToFileUrl(out), tuple(props)))
    except Exception:
        close_quietly(doc)
        raise
    return doc


def reply(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def main():
    try:
        start()
    except Exception as error:
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
        retried = None
        try:
            try:
                doc = convert(job["src"], job["out"], job.get("filterData"))
            except Exception as first:
                retried = "%s: %s (soffice exit=%s)" % (type(first).__name__, first, proc.poll() if proc else None)
                # 연결이 끊겼다(실측: "URP bridge already disposed") — soffice.exe가 살아 있으면 스스로 다시 켜진 것이니
                # 다시 붙기만 하고, 죽었으면 새로 띄운다. 한 번만 다시 해 보고, 그래도 안 되면 실패를 돌려준다
                # (webviewer가 예전 방식으로 바꿔 변환한다).
                try:
                    if proc is not None and proc.poll() is None:
                        connect()
                    else:
                        start()
                except Exception:
                    start()
                doc = convert(job["src"], job["out"], job.get("filterData"))
            reply({"id": job["id"], "ok": True, "ms": int((time.time() - began) * 1000), "retried": retried})
            close_quietly(doc)
        except Exception as error:
            reply({"id": job["id"], "ok": False, "error": "%s: %s" % (type(error).__name__, error), "trace": traceback.format_exc()[-800:]})
        # 문서를 닫은 뒤 soffice가 멀쩡한지 보고, 아니면(실측: 우리 편집기가 저장한 전환·애니메이션 든 pptx를
        # 닫다가 soffice.bin이 죽고 스스로 다시 켜진다 — 10초 남짓) 다시 붙은 뒤에야 "한가함"을 알린다.
        # 그동안 webviewer는 새 변환을 예전 방식으로 돌려, 사용자가 이 회복을 기다리지 않는다.
        recovered = None
        try:
            desktop.getComponents()
        except Exception as error:
            recovered = "%s: %s" % (type(error).__name__, error)
            try:
                if proc is not None and proc.poll() is None:
                    connect()
                else:
                    start()
            except Exception:
                try:
                    start()
                except Exception:
                    pass
        reply({"idle": True, "recovered": recovered})
    stop()


if __name__ == "__main__":
    main()
