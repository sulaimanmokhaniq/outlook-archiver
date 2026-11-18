"""Start the app.

    python OutlookArchiver.py          use classic Outlook on this laptop
    python OutlookArchiver.py --demo   try it with made-up data (any computer)
"""
import sys


def main() -> None:
    from archiver.gui import run

    if "--demo" in sys.argv:
        from archiver.demo_backend import DemoBackend

        run(DemoBackend)
        return

    if sys.platform != "win32":
        sys.exit("This tool needs Windows with classic Outlook. Use --demo to try it elsewhere.")

    from archiver.outlook_backend import OutlookBackend

    run(OutlookBackend)


if __name__ == "__main__":
    main()
