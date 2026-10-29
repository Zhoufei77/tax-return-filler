#!/usr/bin/env python3
"""Developer packaging tool: allowlist, compile, link-check and verify a ZIP.

No upload, installation, customer file access or network calls.
"""
import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
INCLUDE = [
    "SKILL.md", "README.md", "VERSION", "requirements.txt", "THIRD_PARTY_NOTICES.md",
    "scripts/tax_fill.py", "scripts/demo.py", "scripts/statement_adapter.py", "scripts/eit_rules.py", "scripts/vat_rules.py",
    "scripts/tax_workflow.py", "scripts/vat_workflow.py", "scripts/fact_sources.py", "scripts/demo_business.py", "assets/example-template.xlsx",
    "assets/A200000.xlsx", "assets/A200000.json", "assets/A100000.xlsx", "assets/A100000.json",
    "assets/VAT_SMALL_2026.xlsx", "assets/VAT_SMALL_2026.json", "assets/income-statement-quarter.xlsx",
    "assets/income-statement-annual.xlsx", "assets/vat-income.csv",
    "references/runtime.md", "references/job-contract.md", "references/tax-boundaries.md",
    "references/workbuddy-delivery.md", "references/business-workflows.md", "references/eit-2025-adapter.md",
    "references/vat-2026-adapter.md", "references/fact-sources.md", "research/official-evidence.md",
    "research/acceptance-standard.md", "research/verification-summary.md",
]
REQUIRED = ("name", "display_name", "display_name_en", "description", "description_zh", "description_en", "version", "author")


def sha(data):
    return hashlib.sha256(data).hexdigest()


def validate():
    files = {name: (ROOT / name).read_bytes() for name in INCLUDE}
    text = files["SKILL.md"].decode("utf-8")
    header = re.match(r"\A---\n(.*?)\n---\n", text, re.S)
    if not header:
        raise ValueError("missing YAML frontmatter")
    fields = dict(re.findall(r"^([a-z][a-z_-]*):[ \t]*(.*)$", header[1], re.M))
    if any(not fields.get(k, "").strip() for k in REQUIRED):
        raise ValueError("missing WorkBuddy publishing field")
    if fields["name"] != ROOT.name or fields["version"] != files["VERSION"].decode().strip():
        raise ValueError("name or version mismatch")
    for name, data in files.items():
        if len(Path(name).parts) > 2:
            raise ValueError(f"WorkBuddy platform supports only root/subdirectory/file: {name}")
        if name.endswith(".py"):
            compile(data, name, "exec")
        if name.endswith(".md"):
            content = data.decode("utf-8")
            if re.search(r"/Users/[^/\s]+/|/home/[^/\s]+/", content):
                raise ValueError(f"developer path in distribution: {name}")
            for link in re.findall(r"\[[^\]]+\]\(([^)]+)\)", content):
                if re.match(r"https?://|mailto:|#", link):
                    continue
                resolved = ((ROOT / name).parent / link.split("#")[0]).resolve()
                relative = str(resolved.relative_to(ROOT))
                if relative not in files:
                    raise ValueError(f"missing distributed link: {name} -> {relative}")
    return files


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    files = validate()
    out = args.output.expanduser().resolve()
    if out.exists() or out.with_suffix(out.suffix + ".sha256").exists():
        raise FileExistsError("refusing to overwrite an existing distribution")
    out.parent.mkdir(parents=True, exist_ok=True)
    checksums = "".join(f"{sha(data)}  {name}\n" for name, data in sorted(files.items()))
    with zipfile.ZipFile(out, "x", zipfile.ZIP_DEFLATED) as archive:
        for name, data in sorted(files.items()):
            archive.writestr(f"{ROOT.name}/{name}", data)
        archive.writestr(f"{ROOT.name}/MANIFEST.sha256", checksums)
    with zipfile.ZipFile(out) as archive:
        if len(archive.namelist()) != len(files) + 1:
            raise ValueError("unexpected archive entries")
        for name, data in files.items():
            if archive.read(f"{ROOT.name}/{name}") != data:
                raise ValueError(f"archive content mismatch: {name}")
        if archive.read(f"{ROOT.name}/MANIFEST.sha256").decode() != checksums:
            raise ValueError("manifest mismatch")
    digest = sha(out.read_bytes())
    out.with_suffix(out.suffix + ".sha256").write_text(f"{digest}  {out.name}\n", encoding="utf-8")
    print(json.dumps({"zip": str(out), "sha256": digest, "files": len(files) + 1, "bytes": out.stat().st_size,
                      "metadata": "WorkBuddy required fields present", "resources": "verified byte for byte"}, ensure_ascii=False))


if __name__ == "__main__":
    main()
