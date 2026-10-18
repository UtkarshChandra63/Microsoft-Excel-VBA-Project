# Regional Expense Report Automation: Excel VBA Project with UserForm

> One-click consolidation, formatting and totaling of multi-region expense data, driven by an interactive UserForm that launches automatically when the workbook opens.

---

## 1. Short Description / Purpose

Preparing a consolidated expense report from several regional worksheets is repetitive and error-prone when done by hand: you add headers, format each sheet, total each column, then copy every region into a master list.

This project automates that whole workflow in a macro-enabled Excel workbook (`.xlsm`). Four regional worksheets (**East, West, North, South Records**) with different row counts are:

1. given headers and consistent formatting,
2. totaled automatically with a dynamically placed `SUM` formula, and
3. merged into a single master **Yearly Report** sheet,

all from one **Run Report** button on a UserForm. The same form lets the user jump to any sheet through a dropdown or add a new worksheet by name.

**Result on the sample data:** 4 regional sheets (108 expense records in total) are formatted, totaled and consolidated into one Yearly Report showing a grand total of **2,675,745**, in a single click.

---

## 2. Tech Stack

| Technology | How it is used |
|---|---|
| **Microsoft Excel** (macro-enabled workbook, `.xlsm`) | Application and data layer |
| **VBA (Visual Basic for Applications)** | All automation logic |
| **VBA UserForm** (`frmReport`) | GUI with a ComboBox and CommandButtons |
| **Excel macro recorder** | Used to generate the first draft of the header and formatting macros, then cleaned up by hand |
| **Excel object model** | `Worksheets`, `Range`, `Selection`, `ActiveSheet`, `ActiveCell`, `.End(xlDown/xlUp/xlToRight)`, `.Offset`, `.Address` |
| **Excel formulas** | `SUM()` written dynamically from VBA |

**VBA concepts demonstrated:** `For Each` loops, `Do While` loops, `If / And / <>` conditional logic, `Boolean` flags, string concatenation to build formulas, event procedures (`Workbook_Open`, `UserForm_Initialize`, `ComboBox_Change`, `CommandButton_Click`), `InputBox`, and the `AddItem` method.

---

## 3. Data Source

**Origin:** Sample divisional expense records stored in the workbook itself, one worksheet per region. The files `Before_running_the_Excel_VBA_Procedure.xlsx` and `After_running_the_excel_VBA_procedure.xlsx` show the state before and after the macros run.

**Structure of each regional sheet** (6 columns):

| Column | Header (added by VBA) | Content |
|---|---|---|
| A | Division | Region name (East / West / North / South) |
| B | Category | Expense type, e.g. Rent, Software, Advertising, Salaries, Telemarketing |
| C | Jan | January expense |
| D | Feb | February expense |
| E | Mar | March expense |
| F | Total | Total expense for the row (Jan to Mar) |

**Dataset size (varies per sheet, which is why the code uses dynamic ranges):**

| Worksheet | Records | Sheet total |
|---|---|---|
| EAST RECORDS | 22 | 351,405 |
| WEST RECORDS | 15 | 440,390 |
| NORTH RECORDS | 41 | 923,740 |
| SOUTH RECORDS | 30 | 960,210 |
| **YEARLY REPORT** | **108** | **2,675,745** |

**Before:** raw data with no header row and no formatting.
**After:** header row inserted and styled, currency formatting applied, auto-fit columns, a total formula under the Total column on every sheet, and all regions merged into the Yearly Report.

---

## 4. Features / Highlights

### Automatic UserForm on open
- A `Workbook_Open` event shows the **"Welcome to the Report Form"** UserForm as soon as the workbook opens, so the user never has to hunt for macros or buttons.

### Sheet navigator (dropdown)
- The ComboBox is populated at runtime by looping through every worksheet with `AddItem`. It lists *East Records, West Records, North Records, South Records* and *Yearly Report*.
- Choosing an item jumps straight to that worksheet.

### Add Worksheet button
- Prompts for a name through an `InputBox` and creates a new worksheet with that name at the front of the workbook.

### Run Report button (one click does everything)
The `LoopYearlyReport` procedure loops through all worksheets and, for each regional sheet:

1. **Inserts the header row** (`InsertHeaders`): Division, Category, Jan, Feb, Mar, Total.
2. **Formats the sheet** (`FormatHeaders`): bold 12 pt white header text on a tinted theme-colour fill, Currency style on the numeric data, auto-fitted columns.
3. **Totals the Total column** (`AutomateTotalSum`): finds the last filled cell with `End(xlDown)`, reads its address, and writes `=SUM(F2:F<last>)` in the empty cell right below it.
4. **Copies the data** into the **Yearly Report** sheet, pasting each region directly under the previous one by locating the last used row (`Range("A30000").End(xlUp)`).
5. After the loop, **applies the same header, formatting and total** to the Yearly Report itself.

### Built-in safeguards
- **Run-once protection:** a sheet is only processed if cell A1 is not empty and is not already `"Division"`, so clicking Run Report again will not insert duplicate header rows.
- **Excludes the master sheet** from the regional loop, so the Yearly Report is never copied into itself.
- **First-paste flag:** a `FirstTime` Boolean makes the first region paste at the top and every later region paste on the next empty row.

### Scalable by design
- No hard-coded row counts. Ranges are discovered at run time, so sheets with 15 rows or 15,000 rows are handled the same way.
- Modular code: the three core procedures can be run separately or chained together.

---

## 5. How the Project Was Built (Step by Step)

**Step 1: Header insertion and formatting**
Recorded macros with the macro recorder, inserted and formatted headers, tested them, then wrapped them in a `For Each` loop over all worksheets and added conditional logic.

**Step 2: Automated column totals**
Wrote a procedure that uses the Range `Address` property to find the end of the Total column and drops a `SUM` formula just below it, then looped it across every worksheet.

**Step 3: Yearly Report (master list)**
Added a manually created *Yearly Report* sheet. Used VBA to copy each region's data and paste it into the master list, cleaned up the code, added the logic that prevents repeated runs, and tested the final procedure.

**Step 4: UserForm interface**
Built the form, set control properties and formatting, populated the dropdown in `UserForm_Initialize`, wired the ComboBox `Change` event, the Add Worksheet button (with a name prompt), and the Run Report button, then triggered the form from `Workbook_Open`.

---

## 6. Skills Demonstrated

- Process automation and eliminating repetitive manual work
- Excel VBA programming and event-driven design
- Dynamic range handling and formula generation from code
- UserForm design and user experience
- Data consolidation across multiple sources
- Incremental development, testing and code clean-up

---

## 7. How to Use

1. Open the `.xlsm` file and click **Enable Content** if Excel shows a macro security warning.
2. The **Report Form** appears automatically.
3. Pick a worksheet from the dropdown to navigate to it, or click **Add Worksheet** to create a new one.
4. Click **Run Report** to format every regional sheet, add totals, and build the Yearly Report.

---

## 8. Possible Future Improvements

- Add error handling to the Add Worksheet button, for example for blank names, a cancelled `InputBox`, or a duplicate worksheet name.
- Refresh the dropdown after a worksheet is added, so new sheets appear without reopening the form.
- Replace `Select`/`Selection` with direct object references (for example `ws.Range("A1")`) for faster and more robust code.
- Remove duplicate rows: the sample East and North sheets contain repeated records (e.g. Maintenance, Supplies, Software), which currently count twice in the totals.
- Add a regional summary or chart to the Yearly Report, and an option to export it as PDF.
- Use a dynamic column count instead of a fixed `A1:F1` header range.

---

*Built with Excel VBA as a learning project.*
