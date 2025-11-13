# john_is_not_home.py
# Single-run: nudge cursor 1 pixel left if local time is in [06:00,16:00).
# Suitable for Task Scheduler running every 9 minutes.

import datetime
import sys
import os
import time
import traceback


log = r"C:\Users\heure\nudge_run.log"
with open(log, "a") as f:
    f.write(f"----- start {time.ctime()} -----\n")
    try:
        pass  # original code continues below
    except Exception:
        f.write(traceback.format_exc())
        raise
    finally:
        f.write(f"----- end {time.ctime()} -----\n")
try:
    import pyautogui
except Exception as e:
    print(f"missing dependency: {e}", file=sys.stderr)
    raise

# Configuration
ACTIVE_START_HOUR = 6      # inclusive
ACTIVE_END_HOUR = 16       # exclusive
DELTA_X = -1               # move left by 1 pixel
LOG_PATH = os.path.join(os.path.expanduser("~"), "john_is_not_home.log")


def now():
    return datetime.datetime.now()


def in_active_window(dt: datetime.datetime) -> bool:
    return ACTIVE_START_HOUR <= dt.hour < ACTIVE_END_HOUR


def append_log(msg: str):
    ts = now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"{ts}  {msg}\n")
    except Exception:
        # if logging fails, don't crash the nudge
        pass


def nudge_left(delta_x: int):
    try:
        x, y = pyautogui.position()
        new_x = x + delta_x
        # clamp to reasonable integer
        new_x = int(max(-32768, min(32767, new_x)))
        pyautogui.moveTo(new_x, y, duration=0)
        append_log(f"nudged from ({x},{y}) to ({new_x},{y})")
    except Exception as e:
        append_log(f"error moving mouse: {e}")


def main():
    t = now()
    if in_active_window(t):
        nudge_left(DELTA_X)
    else:
        append_log("outside active window; no action")


if __name__ == "__main__":
    main()
