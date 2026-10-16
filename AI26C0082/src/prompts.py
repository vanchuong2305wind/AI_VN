CONVENTIONS = """\
Conventions for interpreting questions:
1. Row labels repeat across sections/groups (e.g. "16 to 19 years" under Total, Men, Women). Always filter by the right
   section/group. If the question names no group (sex, race, ...), use the overall/Total one.
2. Age questions ("youngest", "oldest", "age") are measured at the date the question refers to
   (e.g. "on their admission date" -> admission date minus birth date). Youngest = smallest age.
   If the table has a birth date AND an event date (admission, hire, joining, ...), age always means age at that
   event date (event date minus birth date), even when the question does not mention the date
   (e.g. "the second-oldest person" -> the person with the second-largest value of (event date - birth date)).
3. Sums/totals: if the table already has a Total row/column for what is asked, use it; never double count
   subtotal rows together with their child rows.
4. "Average across all <categories>" = simple mean of every value shown for those categories in that row/column
   (including Total/subtotal columns that appear among them), unless the question excludes some.
5. Values are used exactly as shown in the table (do not convert units, e.g. keep "in thousands" numbers as-is).
6. Text answers: copy the label exactly as written in the table. If the question itself lists names
   (e.g. "rank Agriculture, Manufacturing, and Construction"), use the names exactly as written in the question.
7. "Highest/lowest category" questions consider every row/column shown for that dimension, including aggregate ones,
   unless the question says otherwise.
8. `section` = the heading row a data row falls under (e.g. an industry name, "Expenditure category");
   `group` = the row's direct parent in the table's indentation/hierarchy. "Items directly under X" = rows whose
   group is X. Empty cells mean the value is not available (e.g. "(-8-)", "..") - skip them.
9. When asked which sheet, answer with the sheet name shown in the table header (e.g. 'table_1'), not the table id.
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

JUDGE_SYSTEM = f"""You are a meticulous reviewer. Several solvers answered the same question about spreadsheet tables
and disagree. Each candidate answer below shows how many solvers support it, with the code or reasoning of one
supporter. The majority is not always right: decide which answer is correct (or give the correct one) by checking
the exact cells, the filters used (section / group / label / column) and the question's wording yourself.
Pandas-code solvers compute exactly but may pick the wrong rows or columns; table-reading solvers may make
arithmetic slips.

{CONVENTIONS}
Be brief. Finish with a single final line exactly in the form:
ANSWER: <final answer only, in the format requested by the question>"""

JUDGE_USER = """Tables:
{tables}

Question: {question}

{candidates}"""
