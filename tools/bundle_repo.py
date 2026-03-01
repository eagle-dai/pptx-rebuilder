"""Bundle important repository files into a Markdown document with inline contents."""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
IGNORED_DIR_NAMES = {
    "__pycache__",
    ".git",
    "node_modules",
    "dist",
    "build",
    ".venv",
    "venv",
}
IGNORED_SUFFIXES = {
    ".ico",
    ".svg",
    ".png",
    ".jpg",
    ".jpeg",
    ".pyc",
    ".pyo",
    ".exe",
    ".dll",
    ".so",
    ".dylib",
    ".bin",
    ".class",
    ".pptx",
    ".pdf",
}
IGNORED_FILE_NAMES = {
    "AGENTS.md",
    ".env",
    ".env.example",
    ".gitignore",
    "package-lock.json",
    "uv.lock",
}
DEFAULT_TARGETS = [
    "backend",
    "frontend",
    "README.md",
]
DEFAULT_OUTPUT = Path("tools/tmp/repo_bundle.md")
LARGE_FILE_SIZE_BYTES = 20 * 1024


def resolve_targets(raw_targets: list[str]) -> list[Path]:
    """Resolve raw target strings into repo-relative paths."""
    targets = []
    for raw in raw_targets:
        path = (REPO_ROOT / raw).resolve()
        try:
            path.relative_to(REPO_ROOT)
        except ValueError as exc:
            raise ValueError(f"Target {raw!r} escapes repository root") from exc
        if path.parent != REPO_ROOT:
            raise ValueError(f"Target {raw!r} must live in repository root")
        if not path.exists():
            raise FileNotFoundError(f"Target {raw!r} does not exist")
        targets.append(path)
    return targets


def iter_files(path: Path):
    """Yield files under ``path`` (recursively for directories)."""
    if path.is_file():
        if should_skip_file(path):
            return
        yield path
    elif path.is_dir():
        for file_path in path.rglob("*"):
            if not file_path.is_file():
                continue
            if should_skip_file(file_path):
                continue
            yield file_path
    else:
        raise ValueError(f"Unsupported path type: {path}")


def should_skip_file(path: Path) -> bool:
    """Return True when ``path`` should be excluded from the bundle."""
    if any(part in IGNORED_DIR_NAMES for part in path.parts):
        return True

    if "tools" in path.parts:
        return True

    if path.name in IGNORED_FILE_NAMES:
        return True

    if path.suffix.lower() in IGNORED_SUFFIXES:
        return True

    return False


def gather_file_contents(targets: list[Path]) -> list[tuple[Path, str, int]]:
    """Collect file paths and contents for the provided targets."""
    entries: list[tuple[Path, str, int]] = []
    for target in targets:
        for file_path in iter_files(target):
            rel_path = file_path.relative_to(REPO_ROOT)
            try:
                content = file_path.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                print(
                    f"Skipping {rel_path.as_posix()}: not UTF-8 text.",
                    file=sys.stderr,
                )
                continue
            entries.append((rel_path, content, file_path.stat().st_size))
    entries.sort(key=lambda item: item[0].as_posix())
    return entries


def guess_language(path: Path) -> str:
    """Return a Markdown code fence language identifier based on file suffix."""
    suffix = path.suffix.lstrip(".").lower()
    if suffix in {"py"}:
        return "python"
    if suffix in {"md"}:
        return "markdown"
    if suffix in {"json"}:
        return "json"
    if suffix in {"toml"}:
        return "toml"
    if suffix in {"yml", "yaml"}:
        return "yaml"
    if suffix in {"ts"}:
        return "typescript"
    if suffix in {"tsx"}:
        return "tsx"
    if suffix in {"js", "jsx"}:
        return "javascript"
    if suffix in {"html"}:
        return "html"
    if suffix in {"css"}:
        return "css"
    return ""


def select_fence(content: str) -> str:
    """Return a backtick fence long enough to wrap the provided content."""
    longest_sequence = max(
        (len(match.group(0)) for match in re.finditer(r"`+", content)), default=0
    )
    return "`" * max(3, longest_sequence + 1)


def build_markdown(file_entries: list[tuple[Path, str, int]]) -> str:
    """Generate the Markdown document that holds every file inline."""
    header = "# Repository source bundle\n\n"
    sections: list[str] = []
    for rel_path, content, _ in file_entries:
        language = guess_language(rel_path)
        fence = select_fence(content)
        opening_fence = f"{fence}{language}" if language else fence
        section = "\n".join(
            [
                f"## `{rel_path.as_posix()}`",
                "",
                opening_fence,
                content.rstrip(),
                fence,
            ]
        ).strip()
        sections.append(section)
    document = header + "\n\n".join(sections).rstrip() + "\n"

    return document


def format_size(size_bytes: int) -> str:
    """Format size in human-readable units."""
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(size_bytes)
    for unit in units:
        if size < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(size)} {unit}"
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size_bytes} B"


def write_markdown(content: str, output_path: Path) -> None:
    """Write the Markdown content to disk."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Bundle important repository files into a Markdown document with inline code."
        )
    )
    parser.add_argument(
        "targets",
        nargs="*",
        default=DEFAULT_TARGETS,
        help=(
            "Root-level paths to include in the archive "
            "(default: backend, frontend, README.md)"
        ),
    )
    parser.add_argument(
        "-o",
        "--output",
        default=str(DEFAULT_OUTPUT),
        help="Output Markdown file path (default: tools/tmp/repo_bundle.md)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    targets = resolve_targets(args.targets)
    file_entries = gather_file_contents(targets)
    markdown = build_markdown(file_entries)
    output_path = (REPO_ROOT / args.output).resolve()
    write_markdown(markdown, output_path)

    try:
        output_display = output_path.relative_to(REPO_ROOT).as_posix()
    except ValueError:
        output_display = str(output_path)

    print("Included files:")
    for rel_path, _, size_bytes in file_entries:
        if size_bytes >= LARGE_FILE_SIZE_BYTES:
            print(f"- {rel_path.as_posix()} ({format_size(size_bytes)})")
        else:
            print(f"- {rel_path.as_posix()}")
    print(f"Bundle written to: {output_display}")


if __name__ == "__main__":
    main()
