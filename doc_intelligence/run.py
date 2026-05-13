"""
run.py — Launch the doc_intelligence web server.

# Required commands at runtime:
#   python doc_intelligence/run.py
#   (or via launcher: start.bat at project root)
#
# Prerequisites:
#   - pip install -r requirements_doc_intelligence.txt
#   - Frontend bundle built once: `python doc_intelligence/build.py`
#
# Behavior:
#   - Starts Flask + SocketIO on http://127.0.0.1:5000
#   - Opens the default browser after ~1s delay
#   - Shuts down (os._exit) when all browser clients disconnect
#     and a 5-second grace period elapses.
"""
import threading
import webbrowser

from doc_intelligence.web.app import create_app


HOST = "127.0.0.1"
PORT = 5000


def main() -> None:
    app, sio = create_app()
    url = f"http://{HOST}:{PORT}/"
    timer = threading.Timer(1.0, lambda: webbrowser.open(url))
    timer.daemon = True
    timer.start()
    try:
        sio.run(app, host=HOST, port=PORT, allow_unsafe_werkzeug=True)
    except TypeError:
        sio.run(app, host=HOST, port=PORT)


if __name__ == "__main__":
    main()
