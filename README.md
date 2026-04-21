# Prompt Checker

`Prompt Checker` is a small CLI tool that reads a CSV or Excel file of names and prompts, sends each prompt to OpenAI for rubric-based evaluation, and writes the scored results to a CSV or Excel file.

It is designed for the five DepEd Copilot flight categories:

- `safety`
- `compliance`
- `planning`
- `assessment`
- `growth`

## Requirements

- Python 3
- An OpenAI API key

Install dependencies:

```bash
pip install -r requirements.txt
```

Set your API key:

```bash
export OPENAI_API_KEY=your_api_key_here
```

## Usage

Run the program with:

```bash
python app.py --category safety --input examples/safety.xlsx --output result.xlsx
```

Arguments:

- `--category`: one of `safety`, `compliance`, `planning`, `assessment`, `growth`
- `--input`: input `.csv` or `.xlsx` file
- `--output`: output `.csv` or `.xlsx` file
- `--model`: optional OpenAI model override. Default is `gpt-5-mini`
- `--workers`: optional number of rows to assess in parallel. Default scales to the machine, capped at `8`

Examples:

```bash
python app.py --category planning --input examples/planning.xlsx --output planning_result.xlsx
```

```bash
python app.py --category safety --input examples/safety.csv --output result.csv
```

```bash
python app.py --category safety --input examples/safety.csv --output result.csv --workers 4
```

## Input File Format

The input file must contain at least these columns in the first row:

```text
Name | Prompt
```

Example:

```text
Name   | Prompt
Mikka  | Context: ... Objective: ... References: ... Expectations: ...
Kevin  | Make me a lesson plan...
Happy  | Context: ... Objective: ...
```

Supported input formats:

- `.csv`
- `.xlsx`

## Output File Format

The output file keeps the original `Name` and `Prompt` columns, then adds the rubric columns for the selected category, followed by:

- `Total Score`
- `Overall Feedback`

For example, `safety` outputs:

```text
Name | Prompt | Use of the CORE Framework (Context, Objective, References, Expectations) | Effectiveness and Clarity of the Instructional Prompt | Ethical, Privacy-Aligned, and Child-Safe Prompting (Justification) | Total Score | Overall Feedback
```

Supported output formats:

- `.csv`
- `.xlsx`

## How Scoring Works

For each row:

1. The program reads the `Prompt`.
2. It sends the prompt to OpenAI with the rubric for the selected flight.
3. OpenAI returns a score from `1` to `4` for each rubric criterion.
4. The program sums the criterion scores into `Total Score`.
5. The program writes the results into the output file.

Rows are assessed in parallel, so total runtime is usually much lower than strict one-by-one processing.

## Rubric Categories

The built-in categories map to the flight missions as follows:

- `safety`: Flight 1, Foundations of Copilot Chat and Responsible AI Use
- `compliance`: Flight 2, Admin and Compliance
- `planning`: Flight 3, Lesson Planning and Instruction
- `assessment`: Flight 4, Visual Aids, Assessment, Quizzes, and Rubrics
- `growth`: Flight 5, Professional Growth

## Notes and Limitations

- The tool scores only what is present in the `Prompt` column.
- Some rubrics refer to items like justifications, reflections, validations, or attached reference files.
- If those artifacts are not present in the workbook, the evaluator scores conservatively based only on the visible prompt text.
- The program supports both `.csv` and `.xlsx`.
- The `.xlsx` support uses a simple built-in reader and writer tailored for this workflow.
- Runtime is dominated by OpenAI API calls, not file I/O.
- Using a smaller model such as the default `gpt-5-mini` and parallel workers reduces total runtime.

## API Key Handling

The program reads the API key from the `OPENAI_API_KEY` environment variable.

Environment variable example:

```bash
export OPENAI_API_KEY=your_api_key_here
```

## Files

- [app.py](/home/ralampay/Desktop/prompt-checker/app.py): CLI program
- [requirements.txt](/home/ralampay/Desktop/prompt-checker/requirements.txt): Python dependency list
- `examples/*.csv` and `examples/*.xlsx`: sample inputs, if present

## Quick Start

```bash
pip install -r requirements.txt
export OPENAI_API_KEY=your_api_key_here
python app.py --category safety --input examples/safety.csv --output result.csv --workers 4
```
