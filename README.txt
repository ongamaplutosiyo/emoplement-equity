EMPLOYMENT EQUITY WORKBOOK BUILDER
==================================

How to use
1. Double-click index.html (opens in any modern browser; no internet needed).
2. Upload the employee list (.xlsx, .xls or .csv). The workbook can also hold the payroll sheet.
3. Check the column mapping, the notes and the items to fix.
4. Enter the company name, period, EAP province and sector.
5. Click "Create workbook" to download the full EE workbook (all tabs, with formulas).
6. Open it. The figures show straight away, without clicking Enable Editing.

Privacy
The file is read and built inside your browser. Nothing is uploaded, stored or sent anywhere.

Size
Up to 6,000 employees in the period. Reading and building a large list takes about 30 to 60
seconds. The Quatro Cleaning Services list (6,016 rows) builds this way.

Date format
Choose how text dates in the file are written, on the mapping step, under the column list:
Detect from the file (recommended), day/month/year, month/day/year, or year/month/day.
With Detect, the builder looks at the dates: a first number above 12 means day first, a second
number above 12 means month first. The page shows which order it used. If no date settles it,
day/month/year is used and the page says so. Dates such as 2026/08/31 are always read year first.
Excel date values are not affected by this choice.

Terminations
The Terminations tab lists everyone whose termination date falls in the period, with each
heading filterable. It is rebuilt each time you create a workbook.

Who is included
Only people employed during the period are included: engaged on or before the period end, and
not terminated before the period started. Anyone engaged after the period end is left out, and
so is anyone who left before the period, and independent contractors. Each is listed on the
Import Log with the reason.

Columns
Columns are matched by name, so their order does not matter, and spaces are ignored
("Employee Code" and "EmployeeCode" both match). Required: Employee Code, Gender, Race (RACE or
RacialGroup), Date Engaged, Occupational level. Recommended: Surname (or LastName), Full Names
(or FirstName), Termination Date, Perm/Temp, Nationality (or a country code such as ZAF), Disabled
(or Disability), and the remuneration columns (Salary/Wage ... Other (Fixed), Bonus (Variable)
... Other (Variable)). A "FN" column (foreign national flag) and a single "Total" remuneration
column are also understood.

Perm/Temp from a payroll sheet
If the employee list has no Perm/Temp column (or leaves it blank), the builder reads it from a
payroll sheet in the same file that has Employee Code and Perm/Temp columns.

Items to fix
Problems are listed on the Import Log tab (source row, employee, field, problem, action). Fix
them in the source list and rebuild, or in the workbook directly. The Checks tab shows FIX until
everything is resolved. Anyone with no Perm/Temp status is counted as permanent until one is
entered, and the Checks tab flags them.

Reading the EEA12 figures
- Race is placed by gender and race as recorded. An employee whose nationality is blank is
  counted as South African in the race columns, and the Checks sheet still flags the blank
  nationality. Foreign nationals are counted in their own columns.
- Employees with a race and gender but no occupational level are shown in row 17 ("no level").
- Column N, "Not yet specified", counts employees whose race or gender is still missing.
- Read the male and female lines separately. The TOTAL line nets a male surplus against a
  female shortfall, so it can show zero while one gender is still under target.
- A goal typed as 0 is treated as a 0% target. Leave the box empty for no goal.
- Permanent and temporary staff are shown in separate rows.

Remuneration
EEA4 remuneration is only calculated where the file has remuneration columns. The Quatro list
has no remuneration columns, so its EEA4 figures are blank.

Positions
The Positions tab lists every current position in the file, with the number of employees, the
number employed at the period end, and permanent and temporary counts. Each heading has a filter.
The counts are formulas, so they update if the employee sheet changes.

Filters
Every column heading on the "EEA2 & EEA4 INFO" sheet has a filter button. The Import Log and
Positions tabs have filters too.

Font
All text is Arial 10. Titles and section headings are bold Arial 10.

Licence
The calculation uses HyperFormula (https://hyperformula.handsontable.com), licensed under
GPL-3.0. The rest of this builder is supplied with its source files, as the GPL requires.
