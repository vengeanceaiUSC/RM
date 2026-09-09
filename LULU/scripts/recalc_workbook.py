"""Force LibreOffice to recalculate and save an Excel workbook (incl. external links)."""
import os
import subprocess
import sys
import time


def recalc_workbook(path: str, companion_paths=()) -> str:
    """Open path in headless Calc, calculateAll, save. Returns absolute path."""
    path = os.path.abspath(path)
    companions = [os.path.abspath(p) for p in companion_paths]
    tmp_dir = os.path.dirname(path)

    listener = subprocess.Popen(
        [
            "soffice",
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            f"--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ServiceManager",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    time.sleep(2)
    script = f"""
import uno
import sys
from com.sun.star.beans import PropertyValue

def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p

path = sys.argv[1]
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local)
ctx = resolver.resolve(
    "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
url = uno.systemPathToFileUrl(path)
doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True),))
doc.calculateAll()
doc.store()
doc.close(True)
"""
    env = os.environ.copy()
    env["PYTHONPATH"] = "/usr/lib/python3/dist-packages"
    try:
        subprocess.run(
            [sys.executable, "-c", script, path],
            check=True,
            cwd=tmp_dir,
            env=env,
            capture_output=True,
            text=True,
        )
    finally:
        listener.terminate()
        listener.wait(timeout=10)
    return path


if __name__ == "__main__":
    recalc_workbook(sys.argv[1])
