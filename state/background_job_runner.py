
from __future__ import annotations

import json
import math
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path


def read_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def write_json_atomic(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")
    tmp.replace(path)


def newest(folder: Path, patterns: list[str], since_ts: float | None = None) -> str | None:
    files = []
    for pat in patterns:
        files.extend(folder.glob(pat))
    if since_ts is not None:
        files = [p for p in files if p.exists() and p.stat().st_mtime >= since_ts]
    if not files:
        return None
    return str(max(files, key=lambda p: p.stat().st_mtime))


def run_command(cmd: list[str], cwd: Path, log_file: Path, title: str) -> int:
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    with open(log_file, "a", encoding="utf-8", errors="replace") as log:
        log.write("\n\n" + "=" * 72 + "\n")
        log.write(title + "\n")
        log.write("COMMAND: " + " ".join(str(x) for x in cmd) + "\n")
        log.write("=" * 72 + "\n")
        log.flush()
        proc = subprocess.Popen(
            cmd,
            cwd=str(cwd),
            stdout=log,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
        rc = proc.wait()
        log.write(f"\n[{datetime.now().isoformat(timespec='seconds')}] EXIT CODE: {rc}\n")
        log.flush()
        return rc


def main() -> int:
    job_file = Path(sys.argv[1])
    job = read_json(job_file)
    if not job:
        return 2

    base_dir = Path(job["base_dir"])
    results_dir = Path(job["results_dir"])
    state_dir = Path(job["state_dir"])
    log_file = Path(job["log_file"])
    core_script = Path(job["core_script"])
    symbol = str(job["symbol"]).upper()
    started_ts = float(job.get("started_ts") or time.time())

    def update(status: str, **kwargs) -> None:
        latest = read_json(job_file) or job
        latest.update(kwargs)
        latest["status"] = status
        latest["updated_at"] = datetime.now().isoformat(timespec="seconds")
        write_json_atomic(job_file, latest)

    try:
        update("running", phase="optimize", pid=os.getpid())

        opt_cmd = [sys.executable, "-u", str(core_script)] + job["optimize_args"]
        rc = run_command(opt_cmd, base_dir, log_file, f"Optimize Log - {symbol}")
        if rc != 0:
            update("failed", phase="optimize", returncode=rc, error="Optimize failed. Open the log for details.")
            return rc

        opt_json = newest(results_dir, [f"{symbol}_best_strategy_*.json"], since_ts=started_ts)
        opt_xlsx = newest(results_dir, [f"{symbol}_best_strategy_*.xlsx"], since_ts=started_ts)
        if not opt_json:
            update("failed", phase="optimize", error="Optimize finished but no JSON was found in results/.")
            return 3

        update("running", phase="signal", opt_json=opt_json, opt_xlsx=opt_xlsx)

        signal_started_ts = time.time()
        signal_args = list(job["signal_args"]) + ["--params-file", opt_json]
        sig_cmd = [sys.executable, "-u", str(core_script)] + signal_args
        sig_rc = run_command(sig_cmd, base_dir, log_file, f"Signal Log - {symbol}")

        signal_json = None
        signal_xlsx = None
        signal_payload = {}
        sig = {}
        if sig_rc == 0:
            signal_json = newest(results_dir, [f"{symbol}_signal_*.json"], since_ts=signal_started_ts)
            signal_xlsx = newest(results_dir, [f"{symbol}_signal_*.xlsx"], since_ts=signal_started_ts)
            if signal_json:
                signal_payload = read_json(Path(signal_json))
                sig = signal_payload.get("current_signal", {}) or {}

        payload = read_json(Path(opt_json))
        summary = {}
        summary.update(payload.get("summary", {}) or {})
        summary.update(signal_payload.get("summary", {}) or {})

        analysis_data = {
            "symbol": symbol,
            "inputs": job.get("inputs", {}),
            "summary": summary,
            "current_signal": sig,
            "opt_json": opt_json,
            "opt_xlsx": opt_xlsx,
            "signal_json": signal_json,
            "signal_xlsx": signal_xlsx,
            "package": "",
            "saved_plan": "",
        }

        last_analysis_file = state_dir / "last_analysis.json"
        write_json_atomic(last_analysis_file, analysis_data)

        if sig_rc != 0:
            update(
                "done",
                phase="signal_failed",
                warning="Optimize completed, but Signal failed. The result was saved; check the log.",
                signal_returncode=sig_rc,
                analysis=analysis_data,
            )
        else:
            update("done", phase="complete", returncode=0, analysis=analysis_data)

        return 0

    except Exception as exc:
        update("failed", phase="exception", error=repr(exc))
        return 99


if __name__ == "__main__":
    raise SystemExit(main())
