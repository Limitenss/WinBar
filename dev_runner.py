import sys
import subprocess
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


class Reloader(FileSystemEventHandler):
    def __init__(self, script_to_run):
        self.script = script_to_run
        self.process = None
        self.start_process()

    def start_process(self):
        # If an old version is running, kill it
        if self.process:
            self.process.terminate()
            self.process.wait()
            print("\n[Dev Mode] Restarting App...")

        # Start a fresh version of main.py
        self.process = subprocess.Popen([sys.executable, self.script])

    def on_modified(self, event):
        # Only restart if a Python file was changed (ignore logs or hidden files)
        if event.src_path.endswith(".pyw"):
            self.start_process()


if __name__ == "__main__":
    print("[Dev Mode] Watching for changes...")

    # 1. Define the script to run
    target_script = "main.pyw"

    # 2. Set up the Watchdog observer
    event_handler = Reloader(target_script)
    observer = Observer()

    # "." means watch the current directory
    observer.schedule(event_handler, path=".", recursive=False)
    observer.start()

    try:
        # Keep the runner alive
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        if event_handler.process:
            event_handler.process.terminate()

    observer.join()
