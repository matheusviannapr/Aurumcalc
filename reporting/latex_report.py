from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path
from typing import Any

from .export_tables import export_csv
from .figures import ensure_assets_dir
from .latex_sections import (
    build_assumptions_section,
    build_cover,
    build_end_document,
    build_energy_bill_section,
    build_hourly_balance_section,
    build_load_calibration_section,
    build_objective_section,
    build_preamble,
    build_project_identification,
    build_pv_generation_section,
    build_table_of_contents,
)
from .latex_utils import safe_get, warning_box


def validate_report_context(report_context: dict[str, Any]) -> list[str]:
    warnings: list[str] = []
    required_paths = ["project.client_name", "utility.company", "bill.energy_total_kwh", "pv_generation.annual_kwh"]
    for path in required_paths:
        if safe_get(report_context, path, None) in (None, "não informado"):
            warnings.append(f"Campo obrigatório ausente: {path}")
    return warnings


def generate_latex_report(report_context: dict[str, Any], output_dir: str) -> str:
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    ensure_assets_dir(str(out))
    (out / "anexos").mkdir(exist_ok=True)

    warnings = validate_report_context(report_context)
    report_context.setdefault("warnings", []).extend(warnings)

    sections = [
        build_preamble(report_context),
        build_cover(report_context),
        build_table_of_contents(),
        build_project_identification(report_context),
        build_objective_section(report_context),
        build_assumptions_section(report_context),
        build_energy_bill_section(report_context),
        build_load_calibration_section(report_context),
        build_pv_generation_section(report_context),
        build_hourly_balance_section(report_context),
    ]
    if report_context.get("warnings"):
        sections.append("\\section{Limitações e avisos}" + "\n".join(warning_box(w) for w in report_context["warnings"]))
    sections.append(build_end_document())
    tex_content = "\n\n".join(sections)

    client_slug = str(safe_get(report_context, "project.client_name", "cliente")).lower().replace(" ", "_")
    tex_path = out / f"memoria_calculo_{client_slug}.tex"
    tex_path.write_text(tex_content, encoding="utf-8")
    (out / "report_context.json").write_text(json.dumps(report_context, ensure_ascii=False, indent=2), encoding="utf-8")

    export_csv(out / "anexos" / "carga_horaria.csv", ["hora", "carga_kw"], report_context.get("attachments_data", {}).get("hourly_load", [])[:200])

    pdflatex = shutil.which("pdflatex")
    if pdflatex:
        subprocess.run([pdflatex, "-interaction=nonstopmode", tex_path.name], cwd=out, check=False, capture_output=True)
    return str(tex_path)
