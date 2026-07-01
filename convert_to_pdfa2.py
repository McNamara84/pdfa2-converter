#!/usr/bin/env python3
"""Batch-convert mixed input files to PDF/A-2 PDFs.

The script is intentionally conservative: Python coordinates the workflow,
while format-specific tools do the hard conversion work.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import textwrap
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

try:
    from PIL import Image, ImageDraw, ImageFont, ImageSequence, UnidentifiedImageError
except ImportError:  # pragma: no cover - handled at runtime with a clear message
    Image = None
    ImageDraw = None
    ImageFont = None
    ImageSequence = None
    UnidentifiedImageError = Exception


PDF_EXTENSIONS = {".pdf"}
IMAGE_EXTENSIONS = {
    ".bmp",
    ".gif",
    ".jpeg",
    ".jpg",
    ".png",
    ".tif",
    ".tiff",
    ".webp",
}
TEXT_EXTENSIONS = {
    ".csv",
    ".json",
    ".log",
    ".md",
    ".tsv",
    ".txt",
    ".xml",
    ".xsd",
    ".xslt",
    ".yaml",
    ".yml",
}
OFFICE_EXTENSIONS = {
    ".doc",
    ".docm",
    ".docx",
    ".dot",
    ".dotm",
    ".dotx",
    ".fodp",
    ".fods",
    ".fodt",
    ".odp",
    ".ods",
    ".odt",
    ".otp",
    ".ots",
    ".ott",
    ".pot",
    ".potm",
    ".potx",
    ".pps",
    ".ppsm",
    ".ppsx",
    ".ppt",
    ".pptm",
    ".pptx",
    ".rtf",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
    ".xlt",
    ".xltm",
    ".xltx",
}
SUPPORTED_EXTENSIONS = (
    PDF_EXTENSIONS | IMAGE_EXTENSIONS | TEXT_EXTENSIONS | OFFICE_EXTENSIONS
)

ICC_PROFILE_CANDIDATES = (
    "/usr/share/color/icc/colord/sRGB.icc",
    "/usr/share/ghostscript/iccprofiles/srgb.icc",
    "/usr/share/ghostscript/iccprofiles/esrgb.icc",
    "/usr/share/color/icc/sRGB.icc",
)
FONT_CANDIDATES = (
    "/usr/share/fonts/truetype/liberation2/LiberationMono-Regular.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
)


class ConversionError(Exception):
    """Raised when one file cannot be converted."""

    def __init__(self, stage: str, message: str):
        super().__init__(message)
        self.stage = stage


@dataclass(frozen=True)
class Config:
    input_dir: Path
    output_dir: Path
    workers: int
    overwrite: bool
    recursive: bool
    validate: str
    timeout: int
    keep_temp: bool
    dry_run: bool
    icc_profile: Path
    soffice: str
    ghostscript: str
    verapdf: str | None


@dataclass(frozen=True)
class WorkItem:
    source: Path
    relative_source: Path
    output: Path


@dataclass
class ConversionResult:
    source: str
    output: str
    status: str
    stage: str
    message: str
    seconds: str
    validation: str

    def as_row(self) -> dict[str, str]:
        return {
            "source": self.source,
            "output": self.output,
            "status": self.status,
            "stage": self.stage,
            "message": self.message,
            "seconds": self.seconds,
            "validation": self.validation,
        }


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert files from input_files to PDF/A-2 PDFs in output_pdfs.",
    )
    parser.add_argument("--input-dir", default="input_files", type=Path)
    parser.add_argument("--output-dir", default="output_pdfs", type=Path)
    parser.add_argument(
        "--workers",
        default=1,
        type=int,
        help="Number of files to process in parallel. Default: 1.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing PDFs in the output directory.",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="Only process files directly inside input-dir.",
    )
    parser.add_argument(
        "--validate",
        choices=("auto", "never", "required"),
        default="auto",
        help=(
            "auto uses veraPDF when installed, never skips validation, "
            "required fails if veraPDF is unavailable or rejects a file."
        ),
    )
    parser.add_argument(
        "--timeout",
        default=300,
        type=int,
        help="Timeout in seconds per external conversion step. Default: 300.",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep per-file temporary work directories for debugging.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="List planned conversions without creating PDFs.",
    )
    parser.add_argument(
        "--icc-profile",
        default=None,
        type=Path,
        help="Path to an sRGB ICC profile for Ghostscript PDF/A output.",
    )
    return parser.parse_args(argv)


def find_required_tool(names: Iterable[str], label: str) -> str:
    for name in names:
        path = shutil.which(name)
        if path:
            return path
    raise SystemExit(f"Missing required tool: {label}")


def find_icc_profile(explicit: Path | None) -> Path:
    if explicit:
        if explicit.exists():
            return explicit.resolve()
        raise SystemExit(f"ICC profile not found: {explicit}")

    for candidate in ICC_PROFILE_CANDIDATES:
        path = Path(candidate)
        if path.exists():
            return path.resolve()

    raise SystemExit(
        "No sRGB ICC profile found. Pass one explicitly with --icc-profile."
    )


def find_verapdf() -> str | None:
    system_verapdf = shutil.which("verapdf")
    if system_verapdf:
        return system_verapdf

    local_verapdf = Path(__file__).resolve().parent / ".tools" / "verapdf" / "verapdf"
    if local_verapdf.exists() and os.access(local_verapdf, os.X_OK):
        return str(local_verapdf)

    return None


def build_config(args: argparse.Namespace) -> Config:
    input_dir = args.input_dir.resolve()
    output_dir = args.output_dir.resolve()

    if not input_dir.exists():
        raise SystemExit(f"Input directory does not exist: {input_dir}")
    if not input_dir.is_dir():
        raise SystemExit(f"Input path is not a directory: {input_dir}")
    if args.workers < 1:
        raise SystemExit("--workers must be at least 1")

    verapdf = find_verapdf()
    if args.validate == "required" and not verapdf:
        raise SystemExit("Validation is required, but veraPDF is not installed.")

    return Config(
        input_dir=input_dir,
        output_dir=output_dir,
        workers=args.workers,
        overwrite=args.overwrite,
        recursive=not args.no_recursive,
        validate=args.validate,
        timeout=args.timeout,
        keep_temp=args.keep_temp,
        dry_run=args.dry_run,
        icc_profile=find_icc_profile(args.icc_profile),
        soffice=find_required_tool(("soffice", "libreoffice"), "LibreOffice"),
        ghostscript=find_required_tool(("gs",), "Ghostscript"),
        verapdf=verapdf,
    )


def collect_files(input_dir: Path, recursive: bool) -> list[Path]:
    pattern = "**/*" if recursive else "*"
    return sorted(path for path in input_dir.glob(pattern) if path.is_file())


def output_candidate(input_dir: Path, output_dir: Path, source: Path) -> Path:
    relative = source.relative_to(input_dir)
    return output_dir / relative.with_suffix(".pdf")


def build_work_items(config: Config, sources: list[Path]) -> list[WorkItem]:
    first_pass = [output_candidate(config.input_dir, config.output_dir, src) for src in sources]
    counts = Counter(first_pass)
    used: set[Path] = set()
    work_items: list[WorkItem] = []

    for source, candidate in zip(sources, first_pass):
        relative = source.relative_to(config.input_dir)
        output = candidate

        if counts[candidate] > 1:
            suffix = source.suffix.lower().lstrip(".") or "file"
            output = candidate.with_name(f"{candidate.stem}.{suffix}.pdf")

        if output in used:
            digest = hashlib.sha1(str(relative).encode("utf-8")).hexdigest()[:10]
            output = output.with_name(f"{output.stem}.{digest}.pdf")

        used.add(output)
        work_items.append(WorkItem(source=source, relative_source=relative, output=output))

    return work_items


def configure_logging(output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "conversion.log"
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        force=True,
    )
    return log_path


def temporary_workdir(output_dir: Path, keep_temp: bool) -> Path:
    tmp_root = output_dir / ".tmp"
    tmp_root.mkdir(parents=True, exist_ok=True)
    return Path(tempfile.mkdtemp(prefix="job-", dir=tmp_root))


def cleanup_workdir(path: Path, keep_temp: bool) -> None:
    if not keep_temp:
        shutil.rmtree(path, ignore_errors=True)


def run_command(
    args: list[str],
    stage: str,
    timeout: int,
    env: dict[str, str] | None = None,
) -> subprocess.CompletedProcess[str]:
    logging.info("Running %s: %s", stage, " ".join(args))
    try:
        completed = subprocess.run(
            args,
            capture_output=True,
            env=env,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise ConversionError(stage, f"Timeout after {timeout}s") from exc

    if completed.returncode != 0:
        details = "\n".join(
            part.strip() for part in (completed.stdout, completed.stderr) if part.strip()
        )
        raise ConversionError(
            stage,
            f"Command failed with exit code {completed.returncode}: {shorten(details)}",
        )

    return completed


def shorten(message: str, limit: int = 1200) -> str:
    message = " ".join(message.split())
    if len(message) <= limit:
        return message
    return message[: limit - 3] + "..."


def escape_postscript_path(path: Path) -> str:
    text = str(path)
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def write_pdfa_definition(path: Path, icc_profile: Path, title: str) -> None:
    safe_title = title.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    content = f"""%!
/ICCProfile ({escape_postscript_path(icc_profile)}) def
[ /Title ({safe_title})
  /DOCINFO pdfmark
[/_objdef {{icc_PDFA}} /type /stream /OBJ pdfmark
[{{icc_PDFA}} << /N 3 >> /PUT pdfmark
[{{icc_PDFA}} ICCProfile (r) file /PUT pdfmark
[/_objdef {{OutputIntent_PDFA}} /type /dict /OBJ pdfmark
[{{OutputIntent_PDFA}} <<
  /Type /OutputIntent
  /S /GTS_PDFA1
  /DestOutputProfile {{icc_PDFA}}
  /OutputConditionIdentifier (sRGB IEC61966-2.1)
>> /PUT pdfmark
[{{Catalog}} << /OutputIntents [ {{OutputIntent_PDFA}} ] >> /PUT pdfmark
"""
    path.write_text(content, encoding="utf-8")


def ghostscript_to_pdfa2(source_pdf: Path, output_pdf: Path, workdir: Path, config: Config) -> None:
    pdfa_def = workdir / "PDFA_def.ps"
    write_pdfa_definition(pdfa_def, config.icc_profile, output_pdf.stem)
    args = [
        config.ghostscript,
        f"--permit-file-read={config.icc_profile}",
        "-dPDFA=2",
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        "-sDEVICE=pdfwrite",
        "-dPDFACompatibilityPolicy=1",
        "-dEmbedAllFonts=true",
        "-dSubsetFonts=true",
        "-dCompressFonts=true",
        "-sColorConversionStrategy=RGB",
        "-sProcessColorModel=DeviceRGB",
        f"-sOutputFile={output_pdf}",
        str(pdfa_def),
        str(source_pdf),
    ]
    run_command(args, "pdfa2", config.timeout)

    if not output_pdf.exists() or output_pdf.stat().st_size == 0:
        raise ConversionError("pdfa2", "Ghostscript did not create an output PDF.")


def libreoffice_environment(workdir: Path) -> dict[str, str]:
    home = workdir / "lo-home"
    config = workdir / "lo-config"
    cache = workdir / "lo-cache"
    runtime = workdir / "lo-runtime"
    profile = workdir / "lo-profile"

    for path in (home, config, cache, runtime, profile):
        path.mkdir(parents=True, exist_ok=True)
    runtime.chmod(0o700)

    env = os.environ.copy()
    env.update(
        {
            "HOME": str(home),
            "XDG_CONFIG_HOME": str(config),
            "XDG_CACHE_HOME": str(cache),
            "XDG_RUNTIME_DIR": str(runtime),
            "SAL_USE_VCLPLUGIN": "gen",
            "LANG": env.get("LANG", "C.UTF-8"),
        }
    )
    return env


def office_to_pdf(source: Path, target_pdf: Path, workdir: Path, config: Config) -> None:
    outdir = workdir / "office-pdf"
    outdir.mkdir(parents=True, exist_ok=True)
    profile = (workdir / "lo-profile").resolve().as_uri()
    args = [
        config.soffice,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        f"-env:UserInstallation={profile}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(outdir),
        str(source),
    ]
    before = {path.resolve() for path in outdir.glob("*.pdf")}
    run_command(args, "office-to-pdf", config.timeout, env=libreoffice_environment(workdir))

    expected = outdir / f"{source.stem}.pdf"
    produced: Path | None = expected if expected.exists() else None
    if produced is None:
        after = [path for path in outdir.glob("*.pdf") if path.resolve() not in before]
        if len(after) == 1:
            produced = after[0]

    if produced is None or not produced.exists():
        raise ConversionError("office-to-pdf", "LibreOffice did not create a PDF.")

    shutil.move(str(produced), target_pdf)


def image_frame_to_rgb(frame: Image.Image) -> Image.Image:
    frame = frame.copy()
    if frame.mode in ("RGBA", "LA") or (
        frame.mode == "P" and "transparency" in frame.info
    ):
        rgba = frame.convert("RGBA")
        background = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        background.alpha_composite(rgba)
        return background.convert("RGB")
    if frame.mode != "RGB":
        return frame.convert("RGB")
    return frame


def image_to_pdf(source: Path, target_pdf: Path) -> None:
    if Image is None or ImageSequence is None:
        raise ConversionError("image-to-pdf", "Pillow is not installed.")

    try:
        with Image.open(source) as img:
            pages = [image_frame_to_rgb(frame) for frame in ImageSequence.Iterator(img)]
    except UnidentifiedImageError as exc:
        raise ConversionError("image-to-pdf", f"Unsupported image: {exc}") from exc

    if not pages:
        raise ConversionError("image-to-pdf", "Image contains no frames.")

    first, rest = pages[0], pages[1:]
    first.save(
        target_pdf,
        "PDF",
        resolution=300.0,
        save_all=bool(rest),
        append_images=rest,
    )


def read_text_file(source: Path) -> str:
    data = source.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def load_text_font(size: int) -> ImageFont.ImageFont:
    if ImageFont is None:
        raise ConversionError("text-to-pdf", "Pillow is not installed.")

    for candidate in FONT_CANDIDATES:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def split_long_line(line: str, draw: ImageDraw.ImageDraw, font: ImageFont.ImageFont, width: int) -> list[str]:
    if not line:
        return [""]

    chunks: list[str] = []
    current = line.expandtabs(4)
    while current:
        if draw.textlength(current, font=font) <= width:
            chunks.append(current.rstrip())
            break

        low, high = 1, len(current)
        while low < high:
            mid = (low + high + 1) // 2
            if draw.textlength(current[:mid], font=font) <= width:
                low = mid
            else:
                high = mid - 1

        split_at = max(1, low)
        whitespace_at = current.rfind(" ", 0, split_at)
        if whitespace_at > max(20, split_at // 2):
            split_at = whitespace_at + 1

        chunks.append(current[:split_at].rstrip())
        current = current[split_at:].lstrip()

    return chunks


def text_to_pdf(source: Path, target_pdf: Path) -> None:
    if Image is None or ImageDraw is None or ImageFont is None:
        raise ConversionError("text-to-pdf", "Pillow is not installed.")

    text = read_text_file(source)
    page_width, page_height = 1240, 1754  # A4 at roughly 150 dpi.
    margin = 80
    font = load_text_font(18)
    sample = Image.new("RGB", (1, 1), "white")
    draw = ImageDraw.Draw(sample)
    line_height = max(26, int(draw.textbbox((0, 0), "Ag", font=font)[3] * 1.45))
    max_lines = max(1, (page_height - 2 * margin) // line_height)
    usable_width = page_width - 2 * margin

    wrapped_lines: list[str] = []
    for raw_line in text.splitlines() or [""]:
        wrapped_lines.extend(split_long_line(raw_line, draw, font, usable_width))

    pages: list[Image.Image] = []
    header_font = load_text_font(16)
    for page_index in range(0, len(wrapped_lines), max_lines):
        page_lines = wrapped_lines[page_index : page_index + max_lines]
        page = Image.new("RGB", (page_width, page_height), "white")
        page_draw = ImageDraw.Draw(page)
        page_draw.text((margin, 28), source.name, fill=(90, 90, 90), font=header_font)
        y = margin
        for line in page_lines:
            page_draw.text((margin, y), line, fill=(0, 0, 0), font=font)
            y += line_height
        pages.append(page)

    if not pages:
        pages.append(Image.new("RGB", (page_width, page_height), "white"))

    first, rest = pages[0], pages[1:]
    first.save(
        target_pdf,
        "PDF",
        resolution=150.0,
        save_all=bool(rest),
        append_images=rest,
    )


def convert_to_intermediate_pdf(
    source: Path,
    intermediate_pdf: Path,
    workdir: Path,
    config: Config,
) -> None:
    ext = source.suffix.lower()
    if ext in PDF_EXTENSIONS:
        shutil.copy2(source, intermediate_pdf)
    elif ext in IMAGE_EXTENSIONS:
        image_to_pdf(source, intermediate_pdf)
    elif ext in TEXT_EXTENSIONS:
        text_to_pdf(source, intermediate_pdf)
    elif ext in OFFICE_EXTENSIONS:
        office_to_pdf(source, intermediate_pdf, workdir, config)
    else:
        raise ConversionError("detect-format", f"Unsupported file extension: {ext or '<none>'}")

    if not intermediate_pdf.exists() or intermediate_pdf.stat().st_size == 0:
        raise ConversionError("intermediate-pdf", "No intermediate PDF was created.")


def validate_pdf(output_pdf: Path, config: Config) -> str:
    if config.validate == "never":
        return "skipped"
    if not config.verapdf:
        return "skipped: veraPDF not installed"

    args = [config.verapdf, "--format", "text", "--flavour", "2b", str(output_pdf)]
    completed = run_command(args, "validate", config.timeout)
    report = (completed.stdout or completed.stderr or "").strip()
    if "PASS" in report:
        return "passed"

    if config.validate == "required":
        raise ConversionError("validate", shorten(report or "veraPDF did not report PASS."))
    return "failed: " + shorten(report or "veraPDF did not report PASS.")


def convert_one(item: WorkItem, config: Config) -> ConversionResult:
    start = time.monotonic()
    output = item.output
    source_label = str(item.relative_source)
    output_label = str(output.relative_to(config.output_dir))
    workdir = temporary_workdir(config.output_dir, config.keep_temp)

    try:
        if item.source.suffix.lower() not in SUPPORTED_EXTENSIONS:
            return ConversionResult(
                source_label,
                output_label,
                "unsupported",
                "detect-format",
                f"Unsupported file extension: {item.source.suffix or '<none>'}",
                seconds_since(start),
                "not-run",
            )

        if output.exists() and not config.overwrite:
            validation = "not-run"
            message = "Output exists. Use --overwrite to replace it."
            if config.validate != "never":
                validation = validate_pdf(output, config)
                message = "Output exists and was validated. Use --overwrite to replace it."
            return ConversionResult(
                source_label,
                output_label,
                "skipped",
                "existing-output",
                message,
                seconds_since(start),
                validation,
            )

        output.parent.mkdir(parents=True, exist_ok=True)
        intermediate_pdf = workdir / "intermediate.pdf"
        pdfa_tmp = workdir / "pdfa2.pdf"

        convert_to_intermediate_pdf(item.source, intermediate_pdf, workdir, config)
        ghostscript_to_pdfa2(intermediate_pdf, pdfa_tmp, workdir, config)
        validation = validate_pdf(pdfa_tmp, config)

        if validation.startswith("failed") and config.validate == "required":
            raise ConversionError("validate", validation)

        os.replace(pdfa_tmp, output)
        return ConversionResult(
            source_label,
            output_label,
            "ok",
            "done",
            "Converted to PDF/A-2.",
            seconds_since(start),
            validation,
        )
    except ConversionError as exc:
        logging.exception("Failed %s at %s", item.source, exc.stage)
        return ConversionResult(
            source_label,
            output_label,
            "failed",
            exc.stage,
            str(exc),
            seconds_since(start),
            "not-run",
        )
    except Exception as exc:  # pragma: no cover - safety net for batch mode
        logging.exception("Unexpected failure for %s", item.source)
        return ConversionResult(
            source_label,
            output_label,
            "failed",
            "unexpected",
            f"{type(exc).__name__}: {exc}",
            seconds_since(start),
            "not-run",
        )
    finally:
        cleanup_workdir(workdir, config.keep_temp)


def seconds_since(start: float) -> str:
    return f"{time.monotonic() - start:.2f}"


def write_manifest_header(path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=("source", "output", "status", "stage", "message", "seconds", "validation"),
        )
        writer.writeheader()


def append_manifest_row(path: Path, result: ConversionResult) -> None:
    with path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=("source", "output", "status", "stage", "message", "seconds", "validation"),
        )
        writer.writerow(result.as_row())


def print_dry_run(work_items: list[WorkItem], config: Config) -> None:
    for item in work_items:
        marker = "SUPPORTED" if item.source.suffix.lower() in SUPPORTED_EXTENSIONS else "UNSUPPORTED"
        print(f"{marker}: {item.relative_source} -> {item.output.relative_to(config.output_dir)}")


def summarize(results: list[ConversionResult]) -> str:
    counts = Counter(result.status for result in results)
    parts = [f"{status}={count}" for status, count in sorted(counts.items())]
    return ", ".join(parts) if parts else "no files"


def run(config: Config) -> int:
    config.output_dir.mkdir(parents=True, exist_ok=True)
    log_path = configure_logging(config.output_dir)
    manifest_path = config.output_dir / "manifest.csv"

    sources = collect_files(config.input_dir, config.recursive)
    work_items = build_work_items(config, sources)

    if config.dry_run:
        print_dry_run(work_items, config)
        return 0

    write_manifest_header(manifest_path)
    results: list[ConversionResult] = []

    if not work_items:
        print(f"No files found in {config.input_dir}")
        return 0

    print(f"Converting {len(work_items)} file(s) to {config.output_dir}")
    print(f"Manifest: {manifest_path}")
    print(f"Log: {log_path}")
    if config.validate == "auto" and not config.verapdf:
        print("Validation: skipped because veraPDF is not installed")

    if config.workers == 1:
        for index, item in enumerate(work_items, start=1):
            result = convert_one(item, config)
            results.append(result)
            append_manifest_row(manifest_path, result)
            print_progress(index, len(work_items), result)
    else:
        with ThreadPoolExecutor(max_workers=config.workers) as executor:
            futures = {executor.submit(convert_one, item, config): item for item in work_items}
            for index, future in enumerate(as_completed(futures), start=1):
                result = future.result()
                results.append(result)
                append_manifest_row(manifest_path, result)
                print_progress(index, len(work_items), result)

    summary = summarize(results)
    print(f"Done: {summary}")
    logging.info("Done: %s", summary)

    failed = any(result.status in {"failed", "unsupported"} for result in results)
    return 1 if failed else 0


def print_progress(index: int, total: int, result: ConversionResult) -> None:
    wrapped_message = textwrap.shorten(result.message, width=96, placeholder="...")
    print(
        f"[{index}/{total}] {result.status.upper()} "
        f"{result.source} -> {result.output} ({result.stage}: {wrapped_message})"
    )


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    config = build_config(args)
    return run(config)


if __name__ == "__main__":
    raise SystemExit(main())
