from pathlib import Path
import re
import shutil

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "outputs" / "ai_complex_systems_materials"
DOCX = OUT / "人工智能复杂系统调研报告.docx"
FLOWCHART_SRC = OUT / "AI复杂系统经济价值评估模型流程图.png"
FLOWCHART = OUT / "model_flowchart.png"
TEX = OUT / "人工智能复杂系统调研报告.tex"


def latex_escape(text: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    return "".join(replacements.get(ch, ch) for ch in text)


def iter_block_items(doc):
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def table_to_latex(table: Table) -> str:
    rows = [[latex_escape(cell.text.strip().replace("\n", " ")) for cell in row.cells] for row in table.rows]
    if not rows:
        return ""
    col_count = len(rows[0])
    spec = "p{%.2f\\textwidth}" % (0.94 / col_count)
    spec = " ".join([spec] * col_count)
    lines = [
        r"\begin{table}[H]",
        r"\centering",
        r"\small",
        r"\renewcommand{\arraystretch}{1.25}",
        rf"\begin{{tabular}}{{{spec}}}",
        r"\toprule",
        " & ".join(rows[0]) + r" \\",
        r"\midrule",
    ]
    for row in rows[1:]:
        lines.append(" & ".join(row) + r" \\")
    lines.extend([r"\bottomrule", r"\end{tabular}", r"\end{table}", ""])
    return "\n".join(lines)


def paragraph_to_latex(text: str, previous_heading: str | None) -> tuple[str, str | None]:
    clean = text.strip().replace("\u200b", "")
    if not clean:
        return "", previous_heading

    escaped = latex_escape(clean)
    if previous_heading is None:
        return rf"\title{{{escaped}}}", "title"
    if previous_heading == "title":
        return rf"\author{{{escaped}}}", "author"
    if clean == "摘要":
        return r"\begin{abstract}", "abstract"
    if re.match(r"^[一二三四五六七八九十]+、", clean):
        prefix = ""
        if previous_heading == "abstract":
            prefix = r"\end{abstract}" + "\n\n"
        return prefix + rf"\section{{{escaped}}}", "section"
    if clean.startswith("[S") or re.match(r"^\[\w+\]", clean):
        match = re.match(r"^\[([^\]]+)\]\s*(.*)$", clean)
        if match:
            key = re.sub(r"[^A-Za-z0-9:-]", "", match.group(1)) or "ref"
            body = latex_escape(match.group(2))
            return rf"\bibitem{{{key}}} {body}", previous_heading
        return rf"\bibitem{{ref}} {escaped}", previous_heading
    if clean.startswith("AI应用指数 ="):
        formula = r"\[" + "\n" + r"\text{AI应用指数}=0.25\times\text{渗透率}+0.20\times\text{投资强度}+0.20\times\text{数据可得性}+0.20\times\text{利润改善潜力}+0.15\times\text{社会影响权重}" + "\n" + r"\]"
        return formula, previous_heading
    if clean.startswith("资源优化可表示"):
        formula = (
            escaped
            + "\n\n"
            + r"\["
            + "\n"
            + r"\max \sum_i value_i x_i"
            + "\n"
            + r"\quad s.t.\quad"
            + r"\sum_i cost_i x_i \le Budget,\ "
            + r"\sum_i compute_i x_i \le Compute,\ "
            + r"\sum_i talent_i x_i \le Talent"
            + "\n"
            + r"\]"
        )
        return formula, previous_heading
    return escaped + "\n", previous_heading


def build_tex():
    if not DOCX.exists():
        raise FileNotFoundError(DOCX)
    if not FLOWCHART_SRC.exists():
        raise FileNotFoundError(FLOWCHART_SRC)
    shutil.copyfile(FLOWCHART_SRC, FLOWCHART)

    doc = Document(str(DOCX))
    lines = [
        r"\documentclass[UTF8,a4paper,12pt]{ctexart}",
        r"\usepackage[margin=2.5cm]{geometry}",
        r"\usepackage{graphicx}",
        r"\usepackage{booktabs}",
        r"\usepackage{array}",
        r"\usepackage{float}",
        r"\usepackage{hyperref}",
        r"\usepackage{xcolor}",
        r"\hypersetup{colorlinks=true,linkcolor=blue,urlcolor=blue}",
        r"\graphicspath{{./}{charts/}}",
        r"\setlength{\parindent}{2em}",
        r"\setlength{\parskip}{0.35em}",
        "",
    ]

    state = None
    inserted_flowchart = False
    in_refs = False
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue
            converted, state = paragraph_to_latex(text, state)
            if converted:
                if converted.startswith(r"\bibitem") and not in_refs:
                    lines.append(r"\begin{thebibliography}{99}")
                    in_refs = True
                if in_refs and not converted.startswith(r"\bibitem") and not text.startswith("["):
                    lines.append(r"\end{thebibliography}")
                    in_refs = False
                lines.append(converted)
                if converted.startswith(r"\author"):
                    lines.append(r"\date{}")
                    lines.append(r"\begin{document}")
                    lines.append(r"\maketitle")
                if text.startswith("报告采用随机森林思想") and not inserted_flowchart:
                    lines.extend([
                        "",
                        r"\begin{figure}[H]",
                        r"\centering",
                        r"\includegraphics[width=0.98\textwidth]{model_flowchart.png}",
                        r"\caption{人工智能在复杂系统中的经济价值评估模型流程图}",
                        r"\label{fig:model-flowchart}",
                        r"\end{figure}",
                        "",
                    ])
                    inserted_flowchart = True
                if text.startswith("测算结果显示"):
                    lines.extend([
                        r"\begin{figure}[H]",
                        r"\centering",
                        r"\includegraphics[width=0.88\textwidth]{charts/industry_ai_index.png}",
                        r"\caption{AI在复杂系统行业中的应用成熟度测算}",
                        r"\end{figure}",
                    ])
                if text.startswith("资源优化可表示"):
                    lines.extend([
                        r"\begin{figure}[H]",
                        r"\centering",
                        r"\includegraphics[width=0.88\textwidth]{charts/scenario_output_index.png}",
                        r"\caption{2025--2030年AI驱动综合产出指数预测}",
                        r"\end{figure}",
                    ])
                if text.startswith("WEF Future of Jobs"):
                    lines.extend([
                        r"\begin{figure}[H]",
                        r"\centering",
                        r"\includegraphics[width=0.78\textwidth]{charts/jobs_impact.png}",
                        r"\caption{AI与信息处理技术对岗位结构的影响估计}",
                        r"\end{figure}",
                    ])
        elif isinstance(block, Table):
            lines.append(table_to_latex(block))

    if state == "abstract":
        lines.append(r"\end{abstract}")
    if in_refs:
        lines.append(r"\end{thebibliography}")
    lines.append(r"\end{document}")
    TEX.write_text("\n\n".join(lines), encoding="utf-8")
    return TEX


if __name__ == "__main__":
    print(build_tex())
