"""Start the program.

    python OutlookArchiver.py            use classic Outlook on this laptop
    python OutlookArchiver.py --demo     try it with made-up data (any computer)
    python OutlookArchiver.py --no-window   do not open a window; print the address
"""
import sys


def main() -> None:
    from archiver.server import run_app

    demo = "--demo" in sys.argv
    show = "--no-window" not in sys.argv
    port = 0
    if "--port" in sys.argv:
        port = int(sys.argv[sys.argv.index("--port") + 1])

    if demo:
        from archiver.demo_backend import DemoBackend as Backend
    elif sys.platform != "win32":
        sys.exit("This program needs Windows with classic Outlook. Use --demo to try it elsewhere.")
    else:
        from archiver.outlook_backend import OutlookBackend as Backend

    run_app(Backend, show_window=show, port=port)


if __name__ == "__main__":
    main()
