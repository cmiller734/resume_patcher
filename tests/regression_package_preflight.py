#!/usr/bin/env python3
"""Regression checks for ZIP package preflight validation."""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from zipfile import ZipFile


REPO_ROOT = Path(__file__).resolve().parents[1]
MASTER_RESUME = "Caleb Miller Master Resume.docx"
FORBIDDEN_RESUME = "Caleb Miller - New Resume.docx"
PACKAGE_ROOT = "resume_patcher"


def write_replacements(path: Path) -> None:
    path.write_text(
        json.dumps(
            {
                "paragraph_replacements": [
                    {
                        "find": "Technical Support Engineer with 9 years of experience supporting enterprise SaaS platforms, customer-facing application environments, and production business systems. Strong background owning support cases through resolution, developing deep expertise in a SaaS product set, troubleshooting Windows and Linux environments, and partnering with engineering teams to diagnose product issues. Experienced with REST APIs, SQL, logs, Active Directory, Office 365, SSL/TLS, and customer training in B2B software environments. Combines 6 years of enterprise SaaS support with 3 years of software development experience, bringing rare strength to the overlap between customer-facing troubleshooting and engineering-level problem solving.",
                        "replace": "Regression summary validates package preflight without changing resume strategy.",
                    }
                ]
            },
            indent=2,
        ),
        encoding="utf-8",
    )


def build_package(package_dir: Path, zip_path: Path, *, forbidden_base: bool = False, flat: bool = False) -> None:
    project_dir = package_dir if flat else package_dir / PACKAGE_ROOT
    project_dir.mkdir(parents=True, exist_ok=True)

    for name in [
        MASTER_RESUME,
        "resume_patcher.py",
        "resume_preferences.json",
        "requirements.txt",
        "README.md",
        "CHATBOT_POLICY.md",
    ]:
        shutil.copy2(REPO_ROOT / name, project_dir / name)

    write_replacements(project_dir / "replacements.json")
    manifest = json.loads((REPO_ROOT / "resume_package_manifest.json").read_text(encoding="utf-8"))

    if forbidden_base:
        shutil.copy2(REPO_ROOT / MASTER_RESUME, project_dir / FORBIDDEN_RESUME)
        manifest["master_resume_path"] = FORBIDDEN_RESUME
        manifest["required_files"].append(FORBIDDEN_RESUME)

    (project_dir / "resume_package_manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    (package_dir / ".DS_Store").write_text("ignored metadata", encoding="utf-8")
    (package_dir / "._resume_patcher").write_text("ignored metadata", encoding="utf-8")
    (project_dir / ".DS_Store").write_text("ignored metadata", encoding="utf-8")
    (project_dir / "._Caleb Miller Master Resume.docx").write_text("ignored metadata", encoding="utf-8")
    (project_dir / "~$Caleb Miller Master Resume.docx").write_text("ignored lock", encoding="utf-8")
    macosx_dir = package_dir / "__MACOSX" / PACKAGE_ROOT
    macosx_dir.mkdir(parents=True, exist_ok=True)
    (macosx_dir / "._resume_package_manifest.json").write_text("ignored metadata", encoding="utf-8")

    with ZipFile(zip_path, "w") as zf:
        for path in package_dir.rglob("*"):
            zf.write(path, path.relative_to(package_dir))


def run_package(zip_path: Path, out_path: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            sys.executable,
            str(REPO_ROOT / "resume_patcher.py"),
            "--package",
            str(zip_path),
            "--out",
            str(out_path),
        ],
        cwd=REPO_ROOT,
        text=True,
        capture_output=True,
    )


def run_direct(src_path: Path, replacements_path: Path, out_path: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            sys.executable,
            str(REPO_ROOT / "resume_patcher.py"),
            "--src",
            str(src_path),
            "--replacements",
            str(replacements_path),
            "--out",
            str(out_path),
        ],
        cwd=REPO_ROOT,
        text=True,
        capture_output=True,
    )


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)

        valid_dir = tmp_path / "valid"
        valid_dir.mkdir()
        valid_zip = tmp_path / "valid_package.zip"
        valid_out = tmp_path / "valid_tailored.docx"
        build_package(valid_dir, valid_zip)
        valid_result = run_package(valid_zip, valid_out)
        if valid_result.returncode != 0:
            raise AssertionError(f"valid package failed:\nSTDOUT:\n{valid_result.stdout}\nSTDERR:\n{valid_result.stderr}")
        if not valid_out.exists() or valid_out.stat().st_size <= 0:
            raise AssertionError("valid package did not generate a non-empty output DOCX")

        forbidden_dir = tmp_path / "forbidden"
        forbidden_dir.mkdir()
        forbidden_zip = tmp_path / "forbidden_package.zip"
        forbidden_out = tmp_path / "forbidden_tailored.docx"
        build_package(forbidden_dir, forbidden_zip, forbidden_base=True)
        forbidden_result = run_package(forbidden_zip, forbidden_out)
        if forbidden_result.returncode == 0:
            raise AssertionError("forbidden base resume package unexpectedly succeeded")
        combined_output = f"{forbidden_result.stdout}\n{forbidden_result.stderr}"
        if "Forbidden base resume selected" not in combined_output:
            raise AssertionError(f"forbidden package failed for the wrong reason:\n{combined_output}")
        if forbidden_out.exists():
            raise AssertionError("forbidden package generated an output DOCX")

        flat_dir = tmp_path / "flat"
        flat_dir.mkdir()
        flat_zip = tmp_path / "flat_package.zip"
        flat_out = tmp_path / "flat_tailored.docx"
        build_package(flat_dir, flat_zip, flat=True)
        flat_result = run_package(flat_zip, flat_out)
        if flat_result.returncode == 0:
            raise AssertionError("flat package without resume_patcher/ root unexpectedly succeeded")
        flat_output = f"{flat_result.stdout}\n{flat_result.stderr}"
        if "Required package root folder missing" not in flat_output:
            raise AssertionError(f"flat package failed for the wrong reason:\n{flat_output}")
        if flat_out.exists():
            raise AssertionError("flat package generated an output DOCX")

        direct_replacements = tmp_path / "direct_replacements.json"
        write_replacements(direct_replacements)
        direct_forbidden_src = tmp_path / FORBIDDEN_RESUME
        shutil.copy2(REPO_ROOT / MASTER_RESUME, direct_forbidden_src)
        direct_forbidden_out = tmp_path / "direct_forbidden.docx"
        direct_result = run_direct(direct_forbidden_src, direct_replacements, direct_forbidden_out)
        if direct_result.returncode == 0:
            raise AssertionError("direct mode forbidden base resume unexpectedly succeeded")
        direct_output = f"{direct_result.stdout}\n{direct_result.stderr}"
        if "Forbidden base resume selected" not in direct_output:
            raise AssertionError(f"direct forbidden base failed for the wrong reason:\n{direct_output}")

    print("OK: package preflight accepts only the declared master resume and rejects forbidden bases.")


if __name__ == "__main__":
    main()
