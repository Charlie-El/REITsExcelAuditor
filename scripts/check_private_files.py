from __future__ import annotations

import argparse
from pathlib import Path
import re
import sys


DENY_NAME_PARTS = (
    "待审核",
    "自动审核",
    "批量汇总",
    "补全信息",
)

DENY_DIR_NAMES = {
    "build",
    "dist",
    "__pycache__",
    ".pytest_cache",
}

DENY_SUFFIXES = {
    ".spec",
}

SENSITIVE_CONTENT_PATTERNS = (
    re.compile(r"[A-Za-z]:[\\/](?:Users|用户)[\\/][^\\/\r\n]+"),
    re.compile(r"[A-Za-z]:[\\/][^\r\n]*(?:待审核|自动审核|批量汇总|补全信息)[^\r\n]*\.xls\w*", re.IGNORECASE),
)

TEXT_FILE_SUFFIXES = {
    ".json",
    ".md",
    ".ps1",
    ".py",
    ".toml",
    ".txt",
    ".yaml",
    ".yml",
}

TEXT_FILE_NAMES = {
    ".gitignore",
}

ALLOW_RELATIVE_PATHS = {
    Path("examples") / "metadata_template.xlsx",
}


def is_allowed(path: Path, root: Path) -> bool:
    relative = path.relative_to(root)
    return relative in ALLOW_RELATIVE_PATHS


def is_text_file(path: Path) -> bool:
    return path.suffix.lower() in TEXT_FILE_SUFFIXES or path.name in TEXT_FILE_NAMES


def has_sensitive_content(path: Path) -> bool:
    if not is_text_file(path):
        return False
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return False
    return any(pattern.search(text) for pattern in SENSITIVE_CONTENT_PATTERNS)


def find_private_files(root: Path) -> list[Path]:
    findings: list[Path] = []
    for path in root.rglob("*"):
        if ".git" in path.parts:
            continue
        if path.is_dir():
            continue
        if is_allowed(path, root):
            continue
        relative_parts = path.relative_to(root).parts
        if any(part in DENY_DIR_NAMES for part in relative_parts):
            findings.append(path)
            continue
        if path.suffix.lower() in DENY_SUFFIXES:
            findings.append(path)
            continue
        if any(marker in path.name for marker in DENY_NAME_PARTS):
            findings.append(path)
            continue
        if path.name.startswith("~$"):
            findings.append(path)
            continue
        if has_sensitive_content(path):
            findings.append(path)
            continue
    return findings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Scan the release folder for private or generated files.")
    parser.add_argument("root", nargs="?", default=".", help="Project root to scan.")
    args = parser.parse_args(argv)

    root = Path(args.root).resolve()
    findings = find_private_files(root)
    if findings:
        print("Potential private/generated files found:", file=sys.stderr)
        for path in findings:
            print(f"- {path.relative_to(root)}", file=sys.stderr)
        return 1

    print("Private file scan passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
