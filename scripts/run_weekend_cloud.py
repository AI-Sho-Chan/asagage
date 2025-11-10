#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DEFAULT_CLUSTER = REPO / "infra" / "aws" / "ray-cluster.yaml"


def run(cmd: list[str], *, cwd: Path | None = None, dry_run: bool = False) -> None:
    print("[run]", " ".join(cmd))
    if dry_run:
        return
    proc = subprocess.run(cmd, cwd=cwd or REPO)
    if proc.returncode != 0:
        raise SystemExit(proc.returncode)


def ensure_tool(name: str) -> str:
    path = shutil.which(name)
    if path:
        return path
    # fallbacks for common user-local installs
    defaults = []
    if name == "aws":
        defaults.append(Path.home() / "AppData/Roaming/Python/Python313/Scripts/aws.cmd")
    if name == "ray":
        defaults.append(Path.home() / "AppData/Roaming/Python/Python310/Scripts/ray.exe")
    for candidate in defaults:
        if candidate.exists():
            return str(candidate)
    raise SystemExit(f"Required tool '{name}' not found on PATH (tried {defaults or 'PATH only'})")


def run_local(dry_run: bool) -> None:
    cmd = [
        "powershell",
        "-NoLogo",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(REPO / "scripts" / "run_weekend_then_nightly.ps1"),
    ]
    run(cmd, dry_run=dry_run)


def run_aws(cluster: Path, remote_dir: str, dry_run: bool) -> None:
    ensure_tool("aws")
    ensure_tool("ray")

    run(["ray", "up", "-y", str(cluster)], dry_run=dry_run)
    run(["ray", "rsync_up", str(cluster), str(REPO), remote_dir], dry_run=dry_run)
    run(
        [
            "ray",
            "exec",
            str(cluster),
            f"cd {remote_dir} && python scripts/nightly_build_candidates.py "
            "--run-type weekend --plan-profile weekend --headless",
        ],
        dry_run=dry_run,
    )
    run(
        [
            "ray",
            "rsync_down",
            str(cluster),
            f"{remote_dir}/output",
            str(REPO / "output"),
        ],
        dry_run=dry_run,
    )
    run(["ray", "down", "-y", str(cluster)], dry_run=dry_run)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--backend", choices=["local", "aws"], default="local")
    parser.add_argument("--cluster-config", default=str(DEFAULT_CLUSTER))
    parser.add_argument("--remote-dir", default="/home/ec2-user/asagake")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    cluster = Path(args.cluster_config).resolve()
    if args.backend == "local":
        run_local(args.dry_run)
    else:
        if not cluster.exists():
            raise SystemExit(f"Cluster config not found: {cluster}")
        run_aws(cluster, args.remote_dir, args.dry_run)


if __name__ == "__main__":
    main()
