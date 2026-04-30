from __future__ import annotations

import math
from pathlib import Path
from typing import Any, Iterable, Sequence


def latex_escape(value: Any) -> str:
    text = "não informado" if value is None else str(value)
    replacements = {
        "\\": r"\\textbackslash{}",
        "&": r"\\&",
        "%": r"\\%",
        "$": r"\\$",
        "#": r"\\#",
        "_": r"\\_",
        "{": r"\\{",
        "}": r"\\}",
        "~": r"\\textasciitilde{}",
        "^": r"\\textasciicircum{}",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def safe_get(dictionary: dict, path: str, default: Any = "não informado") -> Any:
    cur: Any = dictionary
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return default
        cur = cur[part]
    return default if cur is None else cur


def _is_valid_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not (math.isnan(value) or math.isinf(value))


def format_currency(value: Any) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if _is_valid_number(value) else "não informado"


def format_kw(value: Any) -> str:
    return f"{value:,.2f} kW".replace(",", "X").replace(".", ",").replace("X", ".") if _is_valid_number(value) else "não informado"


def format_kwh(value: Any) -> str:
    return f"{value:,.2f} kWh".replace(",", "X").replace(".", ",").replace("X", ".") if _is_valid_number(value) else "não informado"


def format_percent(value: Any) -> str:
    return f"{value:.2f}\\%" if _is_valid_number(value) else "não informado"


def format_area(value: Any) -> str:
    return f"{value:,.2f} m$^2$".replace(",", "X").replace(".", ",").replace("X", ".") if _is_valid_number(value) else "não informado"


def build_table(headers: Sequence[str], rows: Iterable[Sequence[Any]], caption: str | None = None, label: str | None = None) -> str:
    cols = "l" * len(headers)
    header_line = " & ".join(latex_escape(h) for h in headers)
    row_lines = [" & ".join(latex_escape(c) for c in row) + r" \\" for row in rows]
    meta = ""
    if caption:
        meta += f"\\caption{{{latex_escape(caption)}}}\n"
    if label:
        meta += f"\\label{{{latex_escape(label)}}}\n"
    return (
        "\\begin{table}[H]\n\\centering\n"
        + meta
        + f"\\begin{{tabular}}{{{cols}}}\n\\toprule\n{header_line} \\\\ \n\\midrule\n"
        + "\n".join(row_lines)
        + "\n\\bottomrule\n\\end{tabular}\n\\end{table}"
    )


def build_longtable(headers: Sequence[str], rows: Iterable[Sequence[Any]], caption: str | None = None, label: str | None = None) -> str:
    cols = "l" * len(headers)
    header_line = " & ".join(latex_escape(h) for h in headers)
    row_lines = [" & ".join(latex_escape(c) for c in row) + r" \\" for row in rows]
    cap = f"\\caption{{{latex_escape(caption)}}}\\\n" if caption else ""
    lab = f"\\label{{{latex_escape(label)}}}\\\n" if label else ""
    return f"\\begin{{longtable}}{{{cols}}}\n{cap}{lab}\\toprule\n{header_line} \\\\ \n\\midrule\n" + "\n".join(row_lines) + "\n\\bottomrule\n\\end{longtable}"


def include_figure(path: str | None, caption: str, label: str, width: str = "0.9\\textwidth") -> str:
    if path and Path(path).exists():
        return (
            "\\begin{figure}[H]\n\\centering\n"
            f"\\includegraphics[width={width}]{{{latex_escape(path)}}}\n"
            f"\\caption{{{latex_escape(caption)}}}\n\\label{{{latex_escape(label)}}}\n\\end{{figure}}"
        )
    return technical_note(f"Figura não disponível: {caption}.")


def section_if_available(title: str, content: str, condition: bool) -> str:
    return f"\\section{{{latex_escape(title)}}}\n{content}" if condition else ""


def warning_box(text: str) -> str:
    return f"\\noindent\\fcolorbox{{red}}{{white}}{{\\parbox{{0.96\\linewidth}}{{\\textbf{{Aviso técnico:}} {latex_escape(text)}}}}}"


def technical_note(text: str) -> str:
    return f"\\noindent\\fcolorbox{{blue}}{{white}}{{\\parbox{{0.96\\linewidth}}{{\\textbf{{Nota técnica:}} {latex_escape(text)}}}}}"
