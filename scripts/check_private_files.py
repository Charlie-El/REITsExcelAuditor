from __future__ import annotations

import argparse
from pathlib import Path
import re
import subprocess
import sys


DENY_DIR_NAMES = {
    ".pytest_cache",
    "__pycache__",
    "build",
    "dist",
    "function",
    "REITsExcelAuditor-main",
}

DENY_SUFFIXES = {
    ".exe",
    ".pdf",
    ".spec",
    ".zip",
}

PRIVATE_WORKBOOK_SUFFIXES = {
    ".xls",
    ".xlsm",
    ".xlsx",
}

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
    Path("examples") / "自动审核补全信息表模板.xlsx",
}

ALLOW_RELATIVE_DIRS = {
    Path("examples") / "annual_update_helper_templates",
    Path("standard_templates"),
}

PRIVATE_NAME_PATTERNS = (
    re.compile(r"结果汇总与复核清单"),
    re.compile(r"人工复核清单"),
    re.compile(r"更新计划预览"),
    re.compile(r"输出对比检查"),
    re.compile(r"AI调用记录"),
    re.compile(r"OCR原始识别结果"),
    re.compile(r"未来现金流汇总表"),
    re.compile(r"基金净资产.*折旧.*摊销.*提取表"),
    re.compile(r"自动更新"),
)

SENSITIVE_CONTENT_PATTERNS = (
    re.compile(r"[A-Za-z]:[\\/](?:Users|Documents and Settings)[\\/][^\\/\r\n]+", re.IGNORECASE),
    re.compile(r"[A-Za-z]:[\\/](?!path[\\/])[^\r\n]*\.(?:xls|xlsx|xlsm|pdf|docx)", re.IGNORECASE),
    re.compile(r"(?:api[_-]?key|secret|token)\s*[:=]\s*['\"][A-Za-z0-9_\-]{16,}['\"]", re.IGNORECASE),
)


def is_allowed(path: Path, root: Path) -> bool:
    relative = path.relative_to(root)
    if relative in ALLOW_RELATIVE_PATHS:
        return True
    return any(is_relative_to(relative, allowed_dir) for allowed_dir in ALLOW_RELATIVE_DIRS)


def is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


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


def is_git_ignored(path: Path, root: Path) -> bool:
    try:
        subprocess.run(
            ["git", "-C", str(root), "check-ignore", "-q", "--", str(path.relative_to(root))],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except (FileNotFoundError, ValueError):
        return False
    except subprocess.CalledProcessError:
        return False
    return True


def find_private_files(root: Path, include_ignored: bool = False) -> list[Path]:
    findings: list[Path] = []
    for path in root.rglob("*"):
        if ".git" in path.parts:
            continue
        if path.is_dir():
            continue
        if not include_ignored and is_git_ignored(path, root):
            continue
        if is_allowed(path, root):
            continue

        relative_parts = path.relative_to(root).parts
        suffix = path.suffix.lower()
        if any(part in DENY_DIR_NAMES for part in relative_parts):
            findings.append(path)
            continue
        if suffix in DENY_SUFFIXES or suffix in PRIVATE_WORKBOOK_SUFFIXES:
            findings.append(path)
            continue
        if any(pattern.search(path.name) for pattern in PRIVATE_NAME_PATTERNS):
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
    parser.add_argument(
        "--include-ignored",
        action="store_true",
        help="Also scan files ignored by Git, such as local function samples, dist, and build outputs.",
    )
    args = parser.parse_args(argv)

    root = Path(args.root).resolve()
    findings = find_private_files(root, include_ignored=args.include_ignored)
    if findings:
        print("Potential private/generated files found:", file=sys.stderr)
        for path in findings:
            print(f"- {path.relative_to(root)}", file=sys.stderr)
        return 1

    print("Private file scan passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
