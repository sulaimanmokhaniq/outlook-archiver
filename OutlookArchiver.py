"""Start the program.

    python OutlookArchiver.py            classic Outlook on this laptop, or Microsoft 365 online
    python OutlookArchiver.py --demo     try it with made-up data (any computer); both
                                         connections (classic and online) are simulated
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
        from archiver.demo_backend import DemoBackend
        from archiver.demo_online import DemoOnlineBackend

        engines = {"classic": lambda _s: DemoBackend(), "online": lambda _s: DemoOnlineBackend()}
    else:
        import os

        from archiver.graph_backend import GraphBackend
        from archiver.storage import app_dir

        def online(settings):
            return GraphBackend(client_id=(settings.get("client_id") or "").strip(),
                                tenant=(settings.get("tenant") or "organizations").strip(),
                                cache_path=os.path.join(app_dir(), "signin.bin"))

        online.needs_client_id = True  # the sign-in needs the organization's app registration
        engines = {"online": online}
        if sys.platform == "win32":
            from archiver.outlook_backend import OutlookBackend

            engines = {"classic": lambda _s: OutlookBackend(), "online": online}

    run_app(engines, show_window=show, port=port)


if __name__ == "__main__":
    main()
