# Gap reporting v1.0.0

Persado EMEA SDE team
george.antoniou@persado.com
First revision: 01/02/2023

## Revisions

v1.1: It now works for column names that begin with "Delivered", "Opened", "Clicked" instead of exact match, because the client is constantly changing the column names.

v1.0.1: Fixed "bug" where results got inverted if there were only 2 variants (ctrl, t1).

v1.0: Final version. Heavily improved, refactored.

v0.9.7: Effort shown in hours instead of minutes (rounded down).

v0.9.5: Removed .xl* from accepted files, fixed bug that created report even with 0 inputs.

v0.9.4: Further optimisation, refactoring. Added Portal link validation (must start with https://portal.persado.com/)

v0.9.1: Can't remember.

v0.9: Refactoring, optimisation.

v0.8: Added support for 9 campaigns in the same report instead of 6.

v0.7: Filename generated based on uploaded file and current date. Added more input validation (won't generate a file if the file is uploaded but there is no other input).

v0.6: Added input validation. It only works if all 3 inputs of a set (campaign name, portal link, geno ids) are entered, otherwise throws an error.

v0.5: Works fine. Added counter.

v0.4: Links work. Also works for 6 campaigns only.

v0.3: Works for multiple campaigns. Does not link the campaign cell.

v0.2: Works with values input by the user. Secondary metrics (average lifts + incr revenue) do not work correctly.

v0.1: Works for one campaign and fixed values for campaign name, portal link, geno ids.

** Only tested in Chrome. **

1. Columns withs same name "Average Lift" renamed to "Average Lift Opens", "Average Lift Clicks", "Average Lift Conversions".
2. Geno IDs should be copy/pasted as they are in the Portal downloaded file (Messages as Excel). Any separator between the IDs will work, as long as there is one (new line, comma, tab, whatever). Should be a new line if you copy/paste them directly from the downloaded Excel. It does not matter if the control Geno IDs is entered first or last (it's last in the Portal downloaded Excel file), the program will "correct" it and put it always first for the final results file.
3. Date of results pulled in the "Results" sheet is entered automatically (current date).
4. There are only a "Results" and a "Raw Data" sheets. "Summary" and "Data" have been removed as they are not needed. All calculations happen programmatically behind the scenes.
5. The final results file gets downloaded automatically after the raw data file is loaded in the app, and then the page reloads to prepare for the next set of results.
6. There is no input validation (for now), so if the program breaks, it's probably wrong/missing input.
7. Cannot programmatically format cells (border, colours etc).
8. Works for cases with 3 geno Ids.
