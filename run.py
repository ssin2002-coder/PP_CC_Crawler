"""Doc Intelligence 웹 서버 진입점"""
import webbrowser
import threading

from doc_intelligence.web.app import create_app


def main():
    app, socketio = create_app()
    port = 5000

    threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{port}")).start()

    print(f"Doc Intelligence running at http://localhost:{port}")
    socketio.run(app, host="0.0.0.0", port=port, debug=False)


if __name__ == "__main__":
    main()
