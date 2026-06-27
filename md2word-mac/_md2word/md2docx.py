from __future__ import annotations

import argparse
import hashlib
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from urllib.parse import urlparse
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from postprocess_docx import main as postprocess_main


TEMPLATE = APP_DIR / "template.docx"
DIAGRAM_FILTER = APP_DIR / "diagram-filter.lua"
REMOTE_IMAGE_RE = re.compile(
    r"!\[([^\]]*)\]\(\s*(?:<(?P<angle_url>https?://[^>]+)>|(?P<plain_url>https?://[^\s)]+))"
    r"(?P<title>\s+(?:\"[^\"]*\"|'[^']*'|\([^)]*\)))?\s*\)"
)


def exe_name(name: str) -> str:
    return f"{name}.exe" if os.name == "nt" else name


def find_tool(name: str, *relative_candidates: str) -> str | None:
    for rel in relative_candidates:
        candidate = APP_DIR / rel
        if candidate.exists():
            return str(candidate)
    return shutil.which(name)


def find_pandoc() -> str:
    pandoc = find_tool(
        "pandoc",
        f"tools/pandoc/{exe_name('pandoc')}",
        f"tools/pandoc/bin/{exe_name('pandoc')}",
    )
    if not pandoc:
        raise SystemExit(
            "Pandoc not found. Put pandoc.exe in tools\\pandoc\\ or install Pandoc and add it to PATH."
        )
    return pandoc


def diagram_tools_available() -> bool:
    mermaid_candidates = [f"tools/mermaid/node_modules/.bin/{exe_name('mmdc')}", "tools/mermaid/node_modules/.bin/mmdc"]
    plantuml_candidates = ["tools/plantuml/plantuml"]
    if os.name == "nt":
        mermaid_candidates.extend(["tools/mermaid/mmdc.cmd", "tools/mermaid/node_modules/.bin/mmdc.cmd"])
        plantuml_candidates.extend(["tools/plantuml/plantuml.bat", "tools/plantuml/plantuml.exe"])

    mmdc = find_tool("mmdc", *mermaid_candidates)
    plantuml = find_tool("plantuml", *plantuml_candidates)
    plantuml_jar = APP_DIR / "tools" / "plantuml" / "plantuml.jar"
    return bool(mmdc or plantuml or plantuml_jar.exists())


def find_browser_for_puppeteer() -> str | None:
    candidates = [
        os.environ.get("PUPPETEER_EXECUTABLE_PATH"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return candidate
    return None


def extension_for_remote_image(url: str, content_type: str | None) -> str:
    parsed_suffix = Path(urlparse(url).path).suffix
    guessed = mimetypes.guess_extension((content_type or "").split(";")[0].strip())
    ext = guessed or parsed_suffix or ".img"
    if ext == ".jpe":
        return ".jpg"
    return ext


def download_remote_image(url: str, output_dir: Path, cache: dict[str, Path]) -> Path:
    if url in cache:
        return cache[url]

    ext = extension_for_remote_image(url, None)
    filename = hashlib.sha1(url.encode("utf-8")).hexdigest() + ext
    output = output_dir / filename
    curl = shutil.which("curl.exe") or shutil.which("curl")
    if curl:
        try:
            subprocess.run(
                [
                    curl,
                    "-L",
                    "--fail",
                    "--connect-timeout",
                    "20",
                    "--max-time",
                    "180",
                    "-A",
                    "Mozilla/5.0 md2word/1.0",
                    "-o",
                    str(output),
                    url,
                ],
                check=True,
            )
            cache[url] = output
            return output
        except Exception as exc:
            print(f"[WARNING] curl could not download remote image {url}: {exc}", file=sys.stderr)

    content_type = None
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 md2word/1.0",
            "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
        },
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        data = response.read()
        content_type = response.headers.get("Content-Type")

    better_ext = extension_for_remote_image(url, content_type)
    if better_ext != ext:
        output = output.with_suffix(better_ext)
    output.write_bytes(data)

    cache[url] = output
    return output


def prepare_markdown_with_remote_images(input_md: Path, output_dir: Path) -> Path:
    text = input_md.read_text(encoding="utf-8")
    cache: dict[str, Path] = {}
    changed = False

    def replace(match: re.Match[str]) -> str:
        nonlocal changed
        alt = match.group(1)
        url = match.group("angle_url") or match.group("plain_url")
        title = match.group("title") or ""
        try:
            local_image = download_remote_image(url, output_dir, cache)
        except Exception as exc:
            print(f"[WARNING] Could not pre-download remote image {url}: {exc}", file=sys.stderr)
            return match.group(0)

        changed = True
        return f"![{alt}](<{local_image.as_posix()}>{title})"

    rewritten = REMOTE_IMAGE_RE.sub(replace, text)
    if not changed:
        return input_md

    prepared_md = output_dir / input_md.name
    prepared_md.write_text(rewritten, encoding="utf-8")
    return prepared_md


def convert(input_md: Path, output_docx: Path, *, toc: bool, diagrams: bool) -> None:
    if not TEMPLATE.exists():
        raise SystemExit(f"Template not found: {TEMPLATE}")

    pandoc = find_pandoc()
    output_docx = output_docx.resolve()
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    raw_docx = output_docx.with_suffix(".raw.docx")

    env = os.environ.copy()
    local_path_entries = []
    local_node = APP_DIR / "tools" / "node"
    local_mermaid = APP_DIR / "tools" / "mermaid"
    local_mermaid_bin = APP_DIR / "tools" / "mermaid" / "node_modules" / ".bin"
    for path in (local_node, local_mermaid, local_mermaid_bin):
        if path.exists():
            local_path_entries.append(str(path))
    if local_path_entries:
        env["PATH"] = os.pathsep.join(local_path_entries + [env.get("PATH", "")])

    browser = find_browser_for_puppeteer()
    if browser and not env.get("PUPPETEER_EXECUTABLE_PATH"):
        env["PUPPETEER_EXECUTABLE_PATH"] = browser

    local_plantuml_jar = APP_DIR / "tools" / "plantuml" / "plantuml.jar"
    if local_plantuml_jar.exists() and not env.get("PLANTUML_JAR"):
        env["PLANTUML_JAR"] = str(local_plantuml_jar)

    with tempfile.TemporaryDirectory(prefix="md2word-resources-") as tmp:
        resource_dir = Path(tmp)
        pandoc_input = prepare_markdown_with_remote_images(input_md, resource_dir)
        cmd = [
            pandoc,
            str(pandoc_input),
            "-f",
            "markdown+task_lists",
            "-o",
            str(raw_docx),
            f"--reference-doc={TEMPLATE}",
            "--standalone",
            "--metadata",
            "toc-title=目录",
            f"--resource-path={os.pathsep.join([str(input_md.parent), str(resource_dir), str(APP_DIR)])}",
        ]
        if toc:
            cmd.append("--toc")
        if diagrams:
            if not DIAGRAM_FILTER.exists():
                raise SystemExit(f"Diagram filter not found: {DIAGRAM_FILTER}")
            cmd.append(f"--lua-filter={DIAGRAM_FILTER}")

        subprocess.run(cmd, check=True, cwd=APP_DIR, env=env)

    old_argv = sys.argv[:]
    try:
        sys.argv = ["postprocess_docx.py", str(raw_docx), str(output_docx)]
        postprocess_main()
    finally:
        sys.argv = old_argv

    raw_docx.unlink(missing_ok=True)
    shutil.rmtree(APP_DIR / ".pandoc-diagrams", ignore_errors=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert Markdown to thesis-formatted Word DOCX.")
    parser.add_argument("input", help="Input Markdown file")
    parser.add_argument("-o", "--output", help="Output DOCX file")
    parser.add_argument("--no-toc", action="store_true", help="Do not generate a table of contents")
    parser.add_argument(
        "--diagrams",
        action="store_true",
        help="Render Mermaid/PlantUML code blocks using diagram-filter.lua",
    )
    parser.add_argument(
        "--auto-diagrams",
        action="store_true",
        help="Enable diagram rendering only when Mermaid/PlantUML tools are available",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_md = Path(args.input).resolve()
    if not input_md.exists():
        raise SystemExit(f"Input file not found: {input_md}")

    output_docx = Path(args.output).resolve() if args.output else input_md.with_suffix(".docx")
    diagrams = args.diagrams or (args.auto_diagrams and diagram_tools_available())
    convert(input_md, output_docx, toc=not args.no_toc, diagrams=diagrams)
    print(f"Created: {output_docx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
