import json
import threading
import time
from pathlib import Path
from urllib.request import urlopen

from desktop import main
from server import run_server


LOCAL_URL = "http://127.0.0.1:8765"


def local_server_ready():
    try:
        with urlopen(LOCAL_URL + "/health", timeout=0.5) as response:
            result=json.loads(response.read().decode("utf-8"))
        return result.get("application")=="Saber Accounting"
    except Exception:
        return False


def start_local_server():
    """Start the private per-PC data service invisibly inside the desktop app."""
    if local_server_ready():
        return
    database=Path.home()/"SaberAccounting"/"saber_accounting_v0_7.db"
    database.parent.mkdir(parents=True,exist_ok=True)
    thread=threading.Thread(target=run_server,
        kwargs={"host":"127.0.0.1","port":8765,"database":str(database),"admin_password":"admin"},
        name="SaberLocalDataService",daemon=True)
    thread.start()
    for _ in range(50):
        if local_server_ready(): return
        if not thread.is_alive(): break
        time.sleep(0.1)
    raise RuntimeError("Saber Accounting could not start its local data service. Close any other copy and try again.")


if __name__ == "__main__":
    start_local_server()
    main()

