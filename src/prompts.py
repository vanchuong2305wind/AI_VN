"""Bước ③: các prompt gửi cho LLM.

Quy ước (CONVENTIONS) là các quy tắc tổng quát rút ra từ tập dev, không viết riêng cho câu hỏi nào.
"""

CONVENTIONS = """\
Conventions for interpreting questions:
1. Row labels repeat across sections/groups (e.g. "16 to 19 years" under Total, Men, Women). Always filter by the right
   section/group. If the question names no group (sex, race, ...), use the overall/Total one.
2. Age questions ("youngest", "oldest", "age") are measured at the date the question refers to
   (e.g. "on their admission date" -> admission date minus birth date). Youngest = smallest age.
3. Sums/totals: if the table already has a Total row/column for what is asked, use it; never double count
   subtotal rows together with their child rows.
4. "Average across all <categories>" = simple mean of every value shown for those categories in that row/column
   (including Total/subtotal columns that appear among them), unless the question excludes some.
5. Values are used exactly as shown in the table (do not convert units, e.g. keep "in thousands" numbers as-is).
6. Text answers: copy the label exactly as written in the table. If the question itself lists names
   (e.g. "rank Agriculture, Manufacturing, and Construction"), use the names exactly as written in the question.
7. "Highest/lowest category" questions consider every row/column shown for that dimension, including aggregate ones,
   unless the question says otherwise.
"""

CODE_SYSTEM = f"""You are an expert data analyst. You answer questions about spreadsheet tables by writing Python (pandas) code.

{CONVENTIONS}
Output format:
- First, 1-3 short lines of plan: which table, which rows/columns (by label and column id), which operation.
- Then ONE python code block defining `def solve(tables):` that returns the final answer.
- `tables` is a dict of pandas DataFrames (described below). pandas as pd, numpy as np, math, re, datetime are available.
- Return a plain value: number, string, bool, or a list (for multi-part answers, in the requested order).
- Do NOT round or format numbers unless the question explicitly asks; formatting is done later.
- Match labels with exact equality on the `label`/`section`/`group` columns as shown (strings are already cleaned).
"""

CODE_USER = """Tables (rendered, values as in the spreadsheet):
{tables}

DataFrames passed to solve(tables):
{frames}

Question: {question}"""

REPAIR_USER = """Your code failed or returned an empty/invalid result:
{error}

Fix the code. Reply with the corrected full `def solve(tables):` in one python code block."""

DIRECT_SYSTEM = f"""You are an expert data analyst. Answer the question by reading the tables carefully.

{CONVENTIONS}
Think step by step briefly: locate the exact cells (row label + column), then compute carefully.
Finish with a single final line exactly in the form:
ANSWER: <final answer only, in the format requested by the question>"""

DIRECT_USER = """Tables:
{tables}

Question: {question}"""

JUDGE_SYSTEM = f"""You are a meticulous reviewer. Two solvers answered the same question about spreadsheet tables
and disagree. Decide which answer is correct (or give the correct one) by checking the cells and the question's
wording yourself.

{CONVENTIONS}
Be brief. Finish with a single final line exactly in the form:
ANSWER: <final answer only, in the format requested by the question>"""

JUDGE_USER = """Tables:
{tables}

Question: {question}

Solver A (pandas code):
{code}
Result A: {result_a}

Solver B (reading the table):
{reasoning_b}
Result B: {result_b}"""
