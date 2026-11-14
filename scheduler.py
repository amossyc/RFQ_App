import time
import config
import main

INTERVAL_MINUTES = getattr(config, "CHECK_INTERVAL_MINUTES", 15)
INTERVAL = INTERVAL_MINUTES * 60

def main_loop():
    print(f"RFQ App scheduler started. Checking inbox every {INTERVAL_MINUTES} minutes.")
    print("Keep this window open. Press Ctrl+C (or close the window) to stop.")
    while True:
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n[{stamp}] Checking inbox ...")
        try:
            main.run()
        except KeyboardInterrupt:
            print("\nStopped by user.")
            return
        except Exception as ex:
            print(f"[{stamp}] ERROR: {ex}")
            print("Will retry at the next interval.")
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Next check in {INTERVAL_MINUTES} minutes.")
        try:
            time.sleep(INTERVAL)
        except KeyboardInterrupt:
            print("\nStopped by user.")
            return

if __name__ == "__main__":
    main_loop()
