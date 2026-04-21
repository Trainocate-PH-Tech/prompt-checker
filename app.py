import argparse
import csv
import json
import os
import sys
import textwrap
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from openai import APIConnectionError, APIStatusError, OpenAI

DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-5")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")


@dataclass(frozen=True)
class RubricDefinition:
    title: str
    criteria: list[str]
    rubric_text: str


RUBRICS: dict[str, RubricDefinition] = {
    "safety": RubricDefinition(
        title="Flight 1: Foundations of Copilot Chat and Responsible AI Use",
        criteria=[
            "Use of the CORE Framework (Context, Objective, References, Expectations)",
            "Effectiveness and Clarity of the Instructional Prompt",
            "Ethical, Privacy-Aligned, and Child-Safe Prompting (Justification)",
        ],
        rubric_text=textwrap.dedent(
            """
            Criteria 1: Use of the CORE Framework (Context, Objective, References, Expectations)
            4 = Clearly includes all four CORE components; each is specific, well-developed, easy to identify; references are appropriate; prompt is reusable and instructional.
            3 = Includes all CORE components, but one may lack detail or clarity; references are present but general.
            2 = Includes some CORE elements, but one or more components are missing, vague, or unclear.
            1 = Does not follow the CORE framework; multiple components are missing or unclear.

            Criteria 2: Effectiveness and Clarity of the Instructional Prompt
            4 = Clear, specific, and efficient; expectations for tone, format, instructional level, and purpose are explicit and strong for classroom use.
            3 = Generally clear and instructional; some expectations could be more explicit.
            2 = Vague or incomplete instructions; response may only partially meet the goal.
            1 = Lacks clarity and structure; response would not align with the intended instructional purpose.

            Criteria 3: Ethical, Privacy-Aligned, and Child-Safe Prompting (Justification)
            4 = Clearly explains ethical AI use, privacy protection, and child-safe design with specific decisions.
            3 = Addresses ethics, privacy, and child-safe considerations, but explanations may be general.
            2 = Mentions ethics or safety but explanation is limited or weakly connected.
            1 = Missing, unclear, or does not demonstrate understanding of ethical, privacy-aligned, or child-safe AI prompting.
            """
        ).strip(),
    ),
    "compliance": RubricDefinition(
        title="Flight 2: Admin and Compliance",
        criteria=[
            "Use of the CORE Framework in the Prompt",
            "Effectiveness and Clarity of the Instructional Prompt",
            "Reflection on Copilot Use and Efficiency",
        ],
        rubric_text=textwrap.dedent(
            """
            Criteria 1: Use of the CORE Framework in the Prompt
            4 = Clearly includes all four CORE components and specifically instructs Copilot to analyze uploaded Excel class records and produce both a narrative report and a presentation.
            3 = Includes all CORE components, but one lacks specificity or clarity.
            2 = Includes some CORE elements, but one or more are missing, vague, or unclear.
            1 = Does not follow the CORE framework; multiple components are missing or unclear.

            Criteria 2: Effectiveness and Clarity of the Instructional Prompt
            4 = Clear, specific, and efficient; gives well-defined instructions for tone, content focus, and output formats for administrative and professional use.
            3 = Generally clear and instructional, but some expectations could be more explicit.
            2 = Vague or incomplete instructions; outputs would only partially meet the administrative purpose.
            1 = Lacks clarity and structure; outputs would not align with the intended task.

            Criteria 3: Reflection on Copilot Use and Efficiency
            4 = Clearly explains how Copilot improved efficiency, organization, or time management with specific administrative tasks.
            3 = Explains support for efficiency, but examples may be general.
            2 = Briefly mentions efficiency gains but with limited explanation.
            1 = Missing, unclear, or does not explain how Copilot improved workflow.
            """
        ).strip(),
    ),
    "planning": RubricDefinition(
        title="Flight 3: Lesson Planning and Instruction",
        criteria=[
            "Use of the CORE Framework in the Prompt",
            "Quality and Clarity of the Instructional Prompt",
            "Reflection on Refining AI Output and Lesson Planning Support",
        ],
        rubric_text=textwrap.dedent(
            """
            Criteria 1: Use of the CORE Framework in the Prompt
            4 = Clearly includes all four CORE components and gives precise guidance for generating a lesson plan aligned with the DepEd curriculum, including subject, grade level, and required lesson components.
            3 = Includes all CORE components, but one element lacks specificity or detail.
            2 = Includes some CORE elements, but one or more are missing or unclear.
            1 = Does not follow the CORE framework; multiple components are missing or unclear.

            Criteria 2: Quality and Clarity of the Instructional Prompt
            4 = Clear, specific, and well-structured, with explicit expectations for lesson objectives, activities, assessment, and curriculum alignment.
            3 = Generally clear and instructional but could use more explicit guidance on lesson structure or alignment.
            2 = Vague or incomplete instructions; lesson plan would be only partially aligned or instructional.
            1 = Lacks clarity and structure; lesson plan would not meet instructional or curriculum requirements.

            Criteria 3: Reflection on Refining AI Output and Lesson Planning Support
            4 = Clearly explains how the teacher refined or improved the AI-generated lesson plan and how Copilot supports planning while maintaining professional judgment.
            3 = Explains support for lesson planning, but refinement details may be general.
            2 = Briefly mentions Copilot use but provides limited explanation of refinement or benefit.
            1 = Missing, unclear, or does not demonstrate understanding of how Copilot supports lesson planning.
            """
        ).strip(),
    ),
    "assessment": RubricDefinition(
        title="Flight 4: Visual Aids, Assessment, Quizzes, and Rubrics",
        criteria=[
            "Use of the CORE Framework in the Prompt",
            "Effectiveness and Clarity of the Instructional Prompt",
            "Validation of Accuracy, Appropriateness, and Curriculum Alignment",
        ],
        rubric_text=textwrap.dedent(
            """
            Criteria 1: Use of the CORE Framework in the Prompt
            4 = Clearly includes all four CORE components and specifically guides Copilot to generate an appropriate learning resource or assessment aligned with curriculum standards.
            3 = Includes all CORE components, but one lacks detail or clarity.
            2 = Includes some CORE elements, but one or more are missing, vague, or unclear.
            1 = Does not follow the CORE framework; multiple components are missing or unclear.

            Criteria 2: Effectiveness and Clarity of the Instructional Prompt
            4 = Clear, specific, and well-structured, with explicit expectations for content type, format, difficulty level, and instructional purpose.
            3 = Generally clear and instructional, but some expectations could be more explicit.
            2 = Vague or incomplete instructions; output would only partially meet goals.
            1 = Lacks clarity and structure; output would not meet the intended instructional purpose.

            Criteria 3: Validation of Accuracy, Appropriateness, and Curriculum Alignment
            4 = Clearly explains how accuracy, learner appropriateness, and curriculum alignment were confirmed, using specific review criteria or strategies.
            3 = Addresses accuracy, appropriateness, and alignment, but explanations may be general.
            2 = Mentions one or two validation aspects but provides limited explanation.
            1 = Missing, unclear, or does not demonstrate understanding of how to evaluate AI-generated instructional materials.
            """
        ).strip(),
    ),
    "growth": RubricDefinition(
        title="Flight 5: Professional Growth",
        criteria=[
            "Use of the CORE Framework in the Prompt",
            "Use of Attached Reference Files",
            "Reflection on Copilot's Support for Growth and Organization",
        ],
        rubric_text=textwrap.dedent(
            """
            Criteria 1: Use of the CORE Framework in the Prompt
            4 = Clearly includes all four CORE components and gives clear guidance for generating a professional growth plan tailored to the teacher's role and career goals.
            3 = Includes all CORE components, but one lacks clarity or specificity.
            2 = Includes some CORE elements, but one or more are missing, vague, or unclear.
            1 = Does not follow the CORE framework; multiple components are missing or unclear.

            Criteria 2: Use of Attached Reference Files
            4 = Appropriately attaches and clearly references relevant career-related files, and effectively directs Copilot to use them in shaping the plan.
            3 = Relevant files are attached and referenced, but instructions for how to use them are general or partially unclear.
            2 = Files are attached but weakly referenced or not clearly integrated into the prompt.
            1 = No relevant reference file is attached, or the file is not referenced in the prompt.

            Criteria 3: Reflection on Copilot's Support for Growth and Organization
            4 = Clearly describes how Copilot supported career pathway exploration and organization of professional goals, with specific examples of efficiency or clarity.
            3 = Addresses support for growth planning and organization, but examples may be general.
            2 = Briefly mentions Copilot assistance but gives limited insight into impact.
            1 = Missing, unclear, or does not demonstrate understanding of how Copilot supports professional growth planning.
            """
        ).strip(),
    ),
}

CLIENT = OpenAI(api_key=OPENAI_API_KEY or None)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Score prompt submissions against a flight rubric and export results to CSV or XLSX."
    )
    parser.add_argument(
        "--category",
        required=True,
        choices=sorted(RUBRICS.keys()),
        help="Rubric category to use: safety, compliance, planning, assessment, or growth.",
    )
    parser.add_argument("--input", required=True, help="Path to the input CSV or XLSX file.")
    parser.add_argument("--output", required=True, help="Path to the output CSV or XLSX file.")
    parser.add_argument(
        "--model",
        default=DEFAULT_MODEL,
        help=f"OpenAI model to use. Defaults to {DEFAULT_MODEL}.",
    )
    return parser.parse_args()


def _xlsx_column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _xlsx_escape(value: Any) -> str:
    text = str(value)
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _get_shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    values: list[str] = []
    for item in root.findall("a:si", ns):
        texts = [node.text or "" for node in item.findall(".//a:t", ns)]
        values.append("".join(texts))
    return values


def read_xlsx_rows(path: Path) -> list[list[str]]:
    with ZipFile(path) as zf:
        shared_strings = _get_shared_strings(zf)
        sheet_name = "xl/worksheets/sheet1.xml"
        if sheet_name not in zf.namelist():
            raise ValueError(f"{path} does not contain xl/worksheets/sheet1.xml")
        root = ET.fromstring(zf.read(sheet_name))
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows: list[list[str]] = []
    for row in root.findall(".//a:sheetData/a:row", ns):
        current: list[str] = []
        for cell in row.findall("a:c", ns):
            cell_type = cell.attrib.get("t")
            value = ""
            if cell_type == "inlineStr":
                parts = [node.text or "" for node in cell.findall(".//a:t", ns)]
                value = "".join(parts)
            elif cell_type == "s":
                node = cell.find("a:v", ns)
                if node is not None and node.text is not None:
                    value = shared_strings[int(node.text)]
            else:
                node = cell.find("a:v", ns)
                if node is not None and node.text is not None:
                    value = node.text
            current.append(value)
        rows.append(current)
    return rows


def write_xlsx_rows(path: Path, rows: list[list[Any]]) -> None:
    def make_sheet_xml(data: list[list[Any]]) -> str:
        lines = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            "  <sheetData>",
        ]
        for row_index, row in enumerate(data, start=1):
            lines.append(f'    <row r="{row_index}">')
            for column_index, value in enumerate(row, start=1):
                cell_ref = f"{_xlsx_column_name(column_index)}{row_index}"
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    lines.append(f'      <c r="{cell_ref}"><v>{value}</v></c>')
                else:
                    text = _xlsx_escape("" if value is None else value)
                    lines.append(
                        f'      <c r="{cell_ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
                    )
            lines.append("    </row>")
        lines.extend(["  </sheetData>", "</worksheet>"])
        return "\n".join(lines)

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Results" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
    core = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-04-21T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-04-21T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""
    app = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>"""

    with ZipFile(path, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", make_sheet_xml(rows))
        zf.writestr("docProps/core.xml", core)
        zf.writestr("docProps/app.xml", app)


def read_csv_rows(path: Path) -> list[list[str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return [row for row in csv.reader(handle)]


def write_csv_rows(path: Path, rows: list[list[Any]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerows(rows)


def read_rows(path: Path) -> list[list[str]]:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return read_xlsx_rows(path)
    if suffix == ".csv":
        return read_csv_rows(path)
    raise ValueError(f"Unsupported input format: {path.suffix}. Use .csv or .xlsx.")


def write_rows(path: Path, rows: list[list[Any]]) -> None:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        write_xlsx_rows(path, rows)
        return
    if suffix == ".csv":
        write_csv_rows(path, rows)
        return
    raise ValueError(f"Unsupported output format: {path.suffix}. Use .csv or .xlsx.")


def normalize_input_rows(rows: list[list[str]]) -> list[dict[str, str]]:
    if not rows:
        raise ValueError("Input workbook is empty.")
    header = [cell.strip() for cell in rows[0]]
    if len(header) < 2:
        raise ValueError("Input workbook must contain at least two columns: Name and Prompt.")
    name_index = None
    prompt_index = None
    for idx, name in enumerate(header):
        lowered = name.lower()
        if lowered == "name":
            name_index = idx
        elif lowered == "prompt":
            prompt_index = idx
    if name_index is None or prompt_index is None:
        raise ValueError("Input workbook header must include Name and Prompt columns.")

    normalized: list[dict[str, str]] = []
    for row in rows[1:]:
        if not any(cell.strip() for cell in row):
            continue
        name = row[name_index].strip() if name_index < len(row) else ""
        prompt = row[prompt_index].strip() if prompt_index < len(row) else ""
        if not name and not prompt:
            continue
        normalized.append({"Name": name, "Prompt": prompt})
    return normalized


def build_json_schema(rubric: RubricDefinition) -> dict[str, Any]:
    score_item = {
        "type": "object",
        "properties": {
            "criterion": {"type": "string", "enum": rubric.criteria},
            "score": {"type": "integer", "minimum": 1, "maximum": 4},
            "rationale": {"type": "string"},
        },
        "required": ["criterion", "score", "rationale"],
        "additionalProperties": False,
    }
    return {
        "name": "rubric_assessment",
        "strict": True,
        "schema": {
            "type": "object",
            "properties": {
                "scores": {
                    "type": "array",
                    "items": score_item,
                    "minItems": len(rubric.criteria),
                    "maxItems": len(rubric.criteria),
                },
                "overall_feedback": {"type": "string"},
            },
            "required": ["scores", "overall_feedback"],
            "additionalProperties": False,
        },
    }


def extract_completion_text(completion: Any) -> str:
    choices = getattr(completion, "choices", None) or []
    if not choices:
        raise ValueError("OpenAI response did not contain any choices.")
    message = choices[0].message
    content = message.content
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        texts: list[str] = []
        for item in content:
            item_type = getattr(item, "type", None)
            if item_type == "text":
                texts.append(getattr(item, "text", ""))
        if texts:
            return "".join(texts)
    raise ValueError("OpenAI response did not contain text content.")


def assess_prompt(prompt: str, rubric: RubricDefinition, model: str, category: str) -> dict[str, Any]:
    if not (OPENAI_API_KEY):
        raise RuntimeError(
            "OPENAI_API_KEY is not set. Export OPENAI_API_KEY or set in app.py."
        )

    system_prompt = textwrap.dedent(
        f"""
        You are an expert rubric assessor for DepEd Copilot flight missions.
        Score only against the provided rubric for category "{category}".
        Return valid JSON that matches the required schema.
        Be strict and evidence-based.

        Important limitation:
        The input may contain only the prompt text, not the supporting justification, reflection, validation, or attached files.
        For criteria that depend on missing artifacts, score conservatively based only on evidence visible in the prompt itself and mention that limitation in the rationale.
        """
    ).strip()

    user_prompt = textwrap.dedent(
        f"""
        Rubric title:
        {rubric.title}

        Criteria:
        {json.dumps(rubric.criteria, ensure_ascii=True, indent=2)}

        Rubric descriptors:
        {rubric.rubric_text}

        Submission prompt to assess:
        {prompt}
        """
    ).strip()

    try:
        completion = CLIENT.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={
                "type": "json_schema",
                "json_schema": build_json_schema(rubric),
            },
            timeout=180,
        )
    except APIStatusError as exc:
        body = exc.response.text if exc.response is not None else str(exc)
        raise RuntimeError(f"OpenAI API request failed with status {exc.status_code}: {body}") from exc
    except APIConnectionError as exc:
        raise RuntimeError(f"OpenAI API request failed: {exc}") from exc

    text = extract_completion_text(completion)
    result = json.loads(text)
    score_map = {item["criterion"]: item for item in result["scores"]}

    ordered_scores = []
    for criterion in rubric.criteria:
        item = score_map.get(criterion)
        if item is None:
            raise ValueError(f"Missing score for criterion: {criterion}")
        ordered_scores.append(item)
    result["scores"] = ordered_scores
    return result


def score_rows(
    category: str, entries: list[dict[str, str]], model: str
) -> list[list[Any]]:
    rubric = RUBRICS[category]
    header = ["Name", "Prompt", *rubric.criteria, "Total Score", "Overall Feedback"]
    output_rows: list[list[Any]] = [header]

    for index, entry in enumerate(entries, start=1):
        name = entry["Name"]
        prompt = entry["Prompt"]
        if not prompt:
            output_rows.append([name, prompt, *([""] * len(rubric.criteria)), "", "Prompt is empty."])
            continue
        print(f"Assessing row {index}/{len(entries)}: {name or '(Unnamed)'}", file=sys.stderr)
        assessment = assess_prompt(prompt, rubric, model, category)
        criterion_scores = [item["score"] for item in assessment["scores"]]
        total = sum(criterion_scores)
        output_rows.append(
            [
                name,
                prompt,
                *criterion_scores,
                total,
                assessment["overall_feedback"],
            ]
        )
    return output_rows


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 1

    try:
        input_rows = read_rows(input_path)
        entries = normalize_input_rows(input_rows)
        scored_rows = score_rows(args.category, entries, args.model)
        write_rows(output_path, scored_rows)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote results to {output_path}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
