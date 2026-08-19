from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
CURRENT_BASELINE = ROOT / "validation" / "pre_thumbnail_current_build.json"


def run(*args: str) -> None:
    subprocess.run([sys.executable, *args], cwd=ROOT, check=True)


def main() -> None:
    run("build_expansion.py")
    run(str(SCRIPTS / "audit_search_thumbnails.py"), "--snapshot", str(CURRENT_BASELINE))
    run(str(SCRIPTS / "apply_search_thumbnails.py"))
    run(str(SCRIPTS / "audit_search_thumbnails.py"), "--baseline", str(CURRENT_BASELINE))


if __name__ == "__main__":
    main()
