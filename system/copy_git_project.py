#!/usr/bin/env python3
import argparse
import logging
import shutil
from pathlib import Path
import sys

try:
    from pathspec import PathSpec
except ImportError:
    sys.exit("Missing dependency: pathspec\nInstall it with: pip install pathspec")

log = logging.getLogger(__name__)


def load_gitignore(src_root: Path) -> PathSpec:
    gitignore_file = src_root / ".gitignore"
    if not gitignore_file.is_file():
        sys.exit(f"No .gitignore found in {src_root}; aborting.")
    lines = gitignore_file.read_text(encoding="utf-8").splitlines()
    return PathSpec.from_lines("gitwildmatch", lines)


def copy_project(src_root: Path, dst_root: Path) -> None:
    spec = load_gitignore(src_root)
    git_dirs_skipped = []

    for path in src_root.rglob("*"):
        rel = path.relative_to(src_root)

        # Skip .git directories to prevent nested git repositories
        if ".git" in path.parts:
            if path.is_dir() and path.name == ".git":
                git_dirs_skipped.append(rel)
            continue

        if spec.match_file(rel.as_posix()):
            continue

        target = dst_root / rel
        if path.is_dir():
            target.mkdir(parents=True, exist_ok=True)
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, target)

    if git_dirs_skipped:
        log.info("Skipped .git directories to prevent nested git repositories:")
        for git_dir in git_dirs_skipped:
            log.info("  - %s", git_dir)


def ask_for_path(prompt: str) -> Path:
    while True:
        try:
            user_input = input(prompt).strip('"').strip("'")
            path = Path(user_input).expanduser().resolve()
            if not path.exists():
                print(f"Path does not exist: {path}")
                continue
            return path
        except Exception as e:
            print(f"Invalid path: {e}")


def main():
    parser = argparse.ArgumentParser(
        description="Copy a git project excluding .gitignore files and .git directories"
    )
    parser.add_argument("source", nargs="?", help="Source folder containing .gitignore")
    parser.add_argument("destination", nargs="?", help="Destination folder")

    args = parser.parse_args()

    if args.source:
        src_root = Path(args.source.strip('"')).expanduser().resolve()
    else:
        src_root = ask_for_path("Enter the source folder path: ")

    if args.destination:
        dst_root = Path(args.destination.strip('"')).expanduser().resolve()
    else:
        dst_root = Path(input("Enter the destination folder path: ").strip('"')).resolve()

    if src_root == dst_root:
        sys.exit("Source and destination must be different.")

    if dst_root.exists() and any(dst_root.iterdir()):
        sys.exit(f"Destination already exists and is not empty: {dst_root}")

    log.info("Copying from:\n  %s\nto:\n  %s", src_root, dst_root)
    log.info("Note: .git directories will be skipped to prevent nested git repositories")
    copy_project(src_root, dst_root)
    log.info("✅ Done!")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    main()
