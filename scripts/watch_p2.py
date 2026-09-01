# scripts/watch_p2.py
# Watchdog-монитор для прогона П2 (run_tests.py).
# Каждые INTERVAL секунд проверяет:
#   - рост logs/test_results.log (размер/время модификации);
#   - живость процесса Excel по PID из logs/excel_pid_tests.txt;
#   - признак завершения ("RunAllTests: END" в хвосте лога).
# Печатает heartbeat-строки "[опроса] HH:MM:SS ..." в консоль и дублирует их
# в logs/monitor_p2.log. При отсутствии роста лога в течение STALL_LIMIT
# интервалов фиксирует возможное зависание.
#
# Запуск (из корня проекта l:/PROject/SysW):
#   python scripts/watch_p2.py

import os
import sys
import time
import subprocess
from datetime import datetime

BASE = os.getcwd()
LOG = os.path.join(BASE, "logs", "test_results.log")
PIDF = os.path.join(BASE, "logs", "excel_pid_tests.txt")
MON = os.path.join(BASE, "logs", "monitor_p2.log")

INTERVAL = 180        # период опроса, сек
STALL_LIMIT = 4       # сколько интервалов без роста считать зависанием


def ts():
    return datetime.now().strftime("%H:%M:%S")


def log(msg):
    line = f"[{ts()}] {msg}"
    print(line, flush=True)
    try:
        os.makedirs(os.path.dirname(MON), exist_ok=True)
        with open(MON, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def pid_exists(pid):
    if not pid:
        return False
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", f"PID eq {pid}"],
            text=True, errors="replace", stderr=subprocess.DEVNULL,
        )
        return str(pid) in out
    except Exception:
        return False


def get_state():
    size = 0
    mtime = 0.0
    last = ""
    try:
        st = os.stat(LOG)
        size = st.st_size
        mtime = st.st_mtime
        with open(LOG, "rb") as f:
            raw = f.read()
        text = raw.decode("utf-8", errors="replace")
        lines = [l for l in text.splitlines() if l.strip()]
        last = lines[-1] if lines else ""
    except Exception:
        pass

    pid = None
    try:
        with open(PIDF, "r", encoding="utf-8") as f:
            pid = int(f.read().strip())
    except Exception:
        pass

    alive = pid_exists(pid)
    return size, mtime, last, pid, alive


def main():
    log(f"--- Старт P2-монитора; интервал={INTERVAL}s; лог={os.path.basename(LOG)} ---")
    prev = None  # (size, mtime)
    stall = 0
    while True:
        time.sleep(INTERVAL)
        size, mtime, last, pid, alive = get_state()

        finished = ("RunAllTests: END" in last) or ("END (Total=" in last)
        grown = prev is not None and (mtime > prev[1] + 0.5 or size != prev[0])

        if finished:
            log(f"[опроса] ЗАВЕРШЕНО по логу: {last[:160]}")
            break

        if grown:
            stall = 0
            status = "ок: прогресс"
        else:
            stall += 1
            status = f"НЕТ РОСТА x{stall}"

        heartbeat = (
            f"[опроса] {ts()} {status}: размер={size}B "
            f"excel_pid={pid}({'жив' if alive else 'мёртв'}) last={last[:90]}"
        )
        log(heartbeat)

        if stall >= STALL_LIMIT:
            log(f"[опроса] ВОЗМОЖНОЕ ЗАВИСАНИЕ: лог не растёт {stall} интервалов, "
                f"excel_pid={pid}({'жив' if alive else 'мёртв'})")
            if not alive:
                log("[опроса] Excel-процесс мёртв, лог не растёт -> считаем прогон завершённым.")
                break

        prev = (size, mtime)

    log("[опроса] Остановка P2-монитора.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("[опроса] Прервано пользователем.")
        sys.exit(0)