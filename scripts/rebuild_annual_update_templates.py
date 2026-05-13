from __future__ import annotations

import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
PS1_SCRIPT = Path(__file__).with_suffix(".ps1")


def template_targets() -> dict[str, Path]:
    template_dir = PROJECT_ROOT / "standard_templates" / "annual_update"
    return {
        "property": template_dir / "产权年报提取模板.xlsx",
        "concession": template_dir / "特许经营权年报提取模板.xlsx",
        "future": template_dir / "未来现金流模板.xlsx",
    }


def source_workbooks() -> dict[str, Path]:
    files = sorted(PROJECT_ROOT.glob("*.xlsx"), key=lambda path: path.stat().st_size)
    result: dict[str, Path] = {}
    for path in files:
        size = path.stat().st_size
        if size > 1_000_000:
            result["property"] = path
        elif 200_000 < size < 300_000:
            result["concession"] = path
        elif 30_000 < size < 60_000:
            result["future"] = path
    return result



def main() -> int:
    if sys.platform != "win32":
        print("This template rebuild step currently requires Windows Excel COM.")
        print(f"Run on Windows with Excel installed: {PS1_SCRIPT}")
        return 1
    sources = source_workbooks()
    targets = template_targets()
    missing = [name for name in ("property", "concession", "future") if name not in sources]
    if missing:
        print(f"Missing source workbook(s): {', '.join(missing)}")
        return 1
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(PS1_SCRIPT),
        "-PropertySource",
        str(sources["property"]),
        "-ConcessionSource",
        str(sources["concession"]),
        "-FutureSource",
        str(sources["future"]),
        "-PropertyTarget",
        str(targets["property"]),
        "-ConcessionTarget",
        str(targets["concession"]),
        "-FutureTarget",
        str(targets["future"]),
    ]
    completed = subprocess.run(command, cwd=PROJECT_ROOT)
    return int(completed.returncode)


if __name__ == "__main__":
    raise SystemExit(main())
