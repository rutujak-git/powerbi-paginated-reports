# Power BI Paginated Reports (Report Builder)

Two production-style paginated reports built end-to-end in **Power BI Report Builder**, covering both major data source patterns you'll encounter in real BI work: a traditional SQL-backed report, and a report connected live to a **Fabric semantic model over an XMLA endpoint** — including writing and debugging **MDX** by hand.

Started from Microsoft's official "Power BI Paginated Reports in a Day" playlist, then extended well beyond the tutorial scope with custom parameters, MDX troubleshooting, conditional formatting, data bars, interactivity, and a full publish-to-service workflow.

## Skills Demonstrated

- **Report Builder / RDL** — parameters, cascading parameters, calculated fields, grouping & subtotals, data bars, indicators, expand/collapse interactivity, click-to-sort
- **MDX** — writing and debugging MDX queries against a Fabric semantic model, including `STRTOSET(..., CONSTRAINED)` member unique-name resolution
- **XMLA endpoint connectivity** — connecting Report Builder to a Fabric semantic model via Azure Analysis Services, instead of querying a table directly
- **Azure SQL Database** — data source setup, SQL query datasets, non-interactive service credentials
- **Power BI Service / Fabric** — publishing paginated reports to a workspace, managing data source credentials for unattended (service-side) execution
- **Practical debugging** — pagination/layout bugs, axis-scaling bugs on data bars, parameter type-mismatch errors, and how each was diagnosed and fixed (documented step by step below, not just the fix — the *reasoning*)

| Report | What it shows | Data source | Demonstrates |
|---|---|---|---|
| `Salesperson_Directory.rdl` | Directory/contact-card layout — one photo + contact block per salesperson | Azure SQL Database (`AdventureWorksDW2019`) | Cascading parameters, calculated fields, image binding, service credential setup |
| `Sales_Performance.rdl` | Country-level COVID-19 stats by year, with charts & interactivity | Fabric semantic model, via XMLA endpoint | MDX, XMLA connectivity, data bars, indicators, nested grouping, interactivity |

Both reports are exported as PDF in this repo (`Salesperson_Directory.pdf`, `Sales_Performance.pdf`) so you can see the output without opening Report Builder.

---

## Data Sources

**Salesperson Directory** — built on the `AdventureWorksDW2019` sample database from Microsoft.
- Download: [AdventureWorksDW2019.bak](https://github.com/Microsoft/sql-server-samples/releases/download/adventureworks/AdventureWorksDW2019.bak)
- Full install instructions: [Microsoft Learn — AdventureWorks Sample Databases](https://learn.microsoft.com/en-us/sql/samples/adventureworks-install-configure?view=sql-server-ver17&tabs=ssms)
- Restored to an Azure SQL Database instance, then queried directly from Report Builder

**Sales Performance (COVID-19 Analysis)** — built on top of `covid19_report.pbix`, a Power BI Desktop file containing the underlying COVID-19 data model (fact/dimension tables, star schema). This `.pbix` was published to the Fabric workspace `covid19_project_workspace` as the semantic model `SM_COVID19_Analysis`, which the paginated report then connects to over the XMLA endpoint (see Report 2, Step 1 below) rather than querying a database table directly.

Both source files (`covid19_report.pbix`) are included in this repo for reference.

---

## Report 1: Salesperson Directory

**Goal:** show every salesperson's photo, title, phone, and email in a clean directory layout, filterable by sales territory group and by individual employee.

<img width="1440" height="850" alt="image" src="https://github.com/user-attachments/assets/0446d128-58c4-4425-9b5b-cab809505f9d" />

<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/80eff64a-8733-457d-8a68-3e37045ed569" />

<img width="1440" height="850" alt="image" src="https://github.com/user-attachments/assets/10bc32d9-df24-455a-898a-68f619485584" />

<img width="1440" height="856" alt="image" src="https://github.com/user-attachments/assets/90f90e1f-11cd-4b60-b992-4a66c8b11bdc" />

### Step 1 — Connect to the data source
- Data source: Azure SQL Database, database `AdventureWorksDW2019`
- Connected using SQL authentication in Report Builder

  <img width="1440" height="856" alt="image" src="https://github.com/user-attachments/assets/7293fe5f-967d-4776-9329-195cf7d86c42" />


### Step 2 — Build the main dataset (`dsMain`)
- Wrote a SQL query joining `DimEmployee` and `DimSalesTerritory` to pull: FirstName, LastName, Title, EmailAddress, Phone, EmployeePhoto, SalesTerritoryRegion, SalesTerritoryGroup

<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/99672b0b-9998-447e-9e9d-56d532ed70bc" />


### Step 3 — Add calculated fields
Two calculated fields were added on top of the raw data:

- **`Salesperson`** — combines first and last name in proper case:
  ```
  =StrConv(Fields!FirstName.Value, vbProperCase) & " " & StrConv(Fields!LastName.Value, vbProperCase)
  ```
StrConv is a VB.NET string-conversion function available in Report Builder expressions. It takes a string and a conversion type, and transforms the text accordingly.

vbProperCase is the specific conversion type that capitalizes the first letter of each word, and lowercases the rest — regardless of how the original text was cased.

  <img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/56c51e8e-eb47-4bde-af2b-3ce226d0a032" />
  

- **`GroupRegion`** — combines territory group and region into one label (e.g. "North America:Northwest"):
  ```
  =Fields!SalesTerritoryGroup.Value & ":" & Fields!SalesTerritoryRegion.Value
  ```

  <img width="1440" height="846" alt="image" src="https://github.com/user-attachments/assets/74add9f7-82fd-494d-b530-fe2aef279995" />


### Step 4 — Build the layout
- Dragged `EmployeePhoto` into an image placeholder (bound as a database image, MIME type `image/jpeg`)
- Stacked `Salesperson`, `GroupRegion`, `Title`, `Phone`, and `EmailAddress` textboxes next to the photo, repeated per employee (one block per row)
- 
<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/847395fe-a504-4dab-aee1-a04fa89910bf" />

### Step 5 — Add cascading parameters
Two parameters were added so the report can be filtered:

1. **`SalesTerritoryGroup`** — dropdown list, values pulled from a new dataset `dsSalesTerritoryGroups`:
   ```sql
   SELECT DISTINCT [SalesTerritoryGroup]
   FROM [dbo].[DimSalesTerritory]
   ORDER BY [SalesTerritoryGroup];
   ```
2. **`EmployeeKey`** — dropdown of employees, filtered by whichever `SalesTerritoryGroup` is picked (this is the "cascading" part — pick a group first, and the employee list narrows to just that group). Multi-select was enabled on this parameter.

<img width="1440" height="856" alt="image" src="https://github.com/user-attachments/assets/81571bbe-920d-4fe9-b29e-b7bc15547ae7" />

<img width="1440" height="854" alt="image" src="https://github.com/user-attachments/assets/5363d2ea-d01e-425b-8f04-05b7f9729a99" />

<img width="1440" height="854" alt="image" src="https://github.com/user-attachments/assets/277211b8-dff1-48b9-8890-a59e9dbf82ec" />

<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/70dcdeea-903e-40a3-b7f8-58848fa82caa" />


### Step 6 — Dynamic subtitle
Added an expression so the subtitle line adjusts automatically depending on what's selected:
```
=Parameters!SalesTerritoryGroup.Value & Iif(Parameters!EmployeeKey.Count = CountRows("dsSalesPerson"), "", " - " & Join(Parameters!EmployeeKey.Label, ", "))
```

<img width="1440" height="858" alt="image" src="https://github.com/user-attachments/assets/a870fd0e-bb2e-4fd2-96e1-436db606fb30" />

This shows just the group name if "all employees" are selected, or the group name plus the specific employee names if a subset is picked.

### Step 7 — Fix layout bug (blank pages)
The report body was 7.6in wide — slightly wider than the printable page area — which caused blank pages to appear between content pages. Fixed by resizing the body to 7.5in.

<img width="1440" height="850" alt="image" src="https://github.com/user-attachments/assets/f2cc15cc-ea4d-4da2-848f-6fff2ddb1e57" />


### Step 8 — Publish and fix credentials
Published the report to the `powerbi_workspace` Fabric workspace. After publishing, the Azure SQL connection needed credentials updated in the service — the original `ActiveDirectoryInteractive` login only works when a person is signing in by hand, but paginated reports need to run unattended in the cloud. Fixed by going to:

**Workspace Settings → Manage connections and gateways → (find the SQL connection) → Edit credentials**, and entering SQL/Azure AD credentials that work without a login prompt.

<img width="1440" height="808" alt="image" src="https://github.com/user-attachments/assets/6112553e-31e5-46f7-be67-c5ab66bf0cc8" />

<img width="1440" height="820" alt="image" src="https://github.com/user-attachments/assets/80ef9889-b6d3-4760-bc3e-d3c609ba7dd7" />

<img width="1440" height="806" alt="image" src="https://github.com/user-attachments/assets/2559a060-598e-4b65-9d65-973cb7cafc9d" />

<img width="1440" height="808" alt="image" src="https://github.com/user-attachments/assets/bc1c0e0f-a7e9-4b2b-ae2c-a613abbfb71b" />


### Known limitation
A few managers (e.g. Amy Alberts, Stephen Jiang, Syed Abbas) show **"NA:NA"** for `GroupRegion` instead of a real territory — this is because their source rows don't have a `SalesTerritoryGroup`/`SalesTerritoryRegion` value at all (they're regional managers, not tied to one specific territory). Left as-is since it accurately reflects the source data, rather than masking it.

---

## Report 2: Sales Performance (COVID-19 Analysis)

**Goal:** show COVID-19 case data by country and year, sourced from a Fabric semantic model instead of a plain SQL table — this report demonstrates connecting Report Builder to a semantic model over an **XMLA endpoint**, plus writing MDX queries.

<img width="1440" height="848" alt="image" src="https://github.com/user-attachments/assets/1d63f560-baff-4226-a532-d1c9a13278f5" />

<img width="1440" height="848" alt="image" src="https://github.com/user-attachments/assets/f2416e72-abc3-4e74-8736-572f74f96849" />

<img width="1440" height="850" alt="image" src="https://github.com/user-attachments/assets/4fb624c5-e8b6-425f-96d7-dda61b055c97" />

### Step 1 — Connect via XMLA endpoint
Instead of a normal SQL data source, this report connects to a **Fabric semantic model** called `SM_COVID19_Analysis`, in the workspace `covid19_project_workspace`.

- Connection type: **Azure Analysis Services**
- Server URL: `powerbi://api.powerbi.com/v1.0/myorg/covid19_project_workspace`
- This lets Report Builder query the semantic model directly (using MDX) instead of hitting a raw database table

<img width="1440" height="814" alt="image" src="https://github.com/user-attachments/assets/ecba6de9-0769-47c3-8265-950f9256e6b9" />

<img width="1440" height="854" alt="image" src="https://github.com/user-attachments/assets/ec06b912-a3db-48fe-bad2-c357cca1bb8a" />


### Step 2 — Build the main dataset (`dsMain`) with MDX
Used the MDX Query Designer (drag-and-drop, not hand-written MDX) to pull: `CountryName`, `Year`, `total_new_confirmed`, `case_fatality_rate`, `full_vaccination_rate_pct`, `SevenDayAvgNewConfirmed`

<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/51674d8a-f4ef-485c-a95d-f387f265177f" />

### Step 3 — Fix the Year parameter (the hardest part)
This was the biggest blocker of the project. Here's what happened and how it was fixed:

**The problem:** Report Builder auto-generates MDX filter code that looks like:
```
STRTOSET(@Year, CONSTRAINED)
```
The `CONSTRAINED` part means the `@Year` parameter has to be passed as an *exact MDX member unique name* — like `[dimdate].[Year].&[2020]` — not just the number `2020`. Typing plain values into the parameter threw this error:
```
The restrictions imposed by the CONSTRAINED flag in the STRTOSET function were violated.
```

**The fix:**
1. Built a small helper dataset called `dsAvailableYears` with this MDX query, which asks for both the plain caption AND the full technical unique name:
   ```
   SELECT
   {} ON COLUMNS,
   { [dimdate].[Year].&[2020], [dimdate].[Year].&[2021], [dimdate].[Year].&[2022] }
   DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS
   FROM [Model]
   ```
2. In the `Year` parameter's properties, set **Available Values → Get values from a query**, using `dsAvailableYears`:
   - **Value field** = the `MEMBER_UNIQUE_NAME` column (the technical name, e.g. `[dimdate].[Year].&[2020]`)
   - **Label field** = the `MEMBER_CAPTION` column (the friendly name, e.g. `2020`)
     
<img width="1440" height="860" alt="image" src="https://github.com/user-attachments/assets/b974d610-3dbe-4abe-8507-b00fbc413327" />

<img width="1436" height="848" alt="image" src="https://github.com/user-attachments/assets/9cb969ff-b942-494a-b482-129e9ade669d" />

<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/601b7988-c3eb-4af6-8a9e-0dfc4f72c938" />


This way, the dropdown displays "2020" to the user, but silently passes the correct technical unique name to the query behind the scenes.


### Step 4 — Multi-select years
Turned on **Allow multiple values** for the Year parameter, so users can pick more than one year at once (e.g. 2020 + 2021 + 2022 together).

<img width="1440" height="858" alt="image" src="https://github.com/user-attachments/assets/5d003139-ed3c-44f7-a326-06af676b6e77" />


This broke the subtitle textbox, since `.Label` now returns a list of years instead of just one — fixed with `Join()`:
```
="Year: " & Join(Parameters!dimdateYear.Label, ", ")
```

### Step 5 — Conditional formatting for missing data
Vaccination data is mostly unavailable for early 2020. Added conditional formatting so any cell showing "no data" for `full_vaccination_rate_pct` displays in **red**, making gaps easy to spot at a glance.

<img width="1440" height="856" alt="image" src="https://github.com/user-attachments/assets/10913412-f22f-4b9f-baba-2db402b1fb72" />

<img width="1438" height="856" alt="image" src="https://github.com/user-attachments/assets/c84fab72-e17d-4802-82b8-317dbe94b549" />

### Step 6 — Filter out fully-empty rows
Some tiny territories (e.g. Christmas Island) have *no* data at all across every metric. Rather than removing all rows with any missing data (which would hide real, useful partial data), added a filter that only removes rows where **every** key metric is missing — a "hybrid" approach that keeps genuinely informative rows.

<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/494fa32f-347b-4220-856f-7d4858753e8f" />


### Step 7 — Group by Year → Country, with subtotals
Added nested grouping so the report shows a subtotal per country (e.g. total confirmed cases per country), with individual months underneath. Verified subtotal math manually against the visible rows to make sure the totals were correct.

<img width="1440" height="854" alt="image" src="https://github.com/user-attachments/assets/5440af40-6a39-4a75-ac51-da8f77b0500c" />

<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/37ba136d-05d5-4b29-b1de-ef583171efe9" />

### Step 8 — Data bars (scaled per country)
Added an in-cell data bar to visualize `total_new_confirmed` per row. Initially every bar looked identical (100% full) — this happens because data bars scale automatically per single value by default. Fixed by setting a fixed axis range:
- **Minimum:** 0
- **Maximum:** `=Max(Fields!total_new_confirmed.Value, "CountryName")`

<img width="1440" height="856" alt="image" src="https://github.com/user-attachments/assets/784f736f-3d22-422c-94aa-3dc30f6f367f" />

<img width="1440" height="850" alt="image" src="https://github.com/user-attachments/assets/78509041-9a9e-4133-b60c-ada8ca0f03e2" />

<img width="1440" height="850" alt="image" src="https://github.com/user-attachments/assets/548add39-dfbf-4a21-a55b-666adf4bb566" />


Using `"CountryName"` as the scope means each country's own bars are scaled against that country's own highest value — so small countries still show meaningful variation, instead of every country being squashed flat by comparison to a huge outlier like Brazil.


### Step 9 — Repeating header row across pages
The column header row (Year / total new confirmed / etc.) was only showing on page 1. Fixed by setting `RepeatOnNewPage = True` on the correct **Row Groups** static member — note: this same setting exists on **Column Groups** too, but setting it there throws an error. It only works on the Row Groups side.

<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/cdc6cb4d-bf6a-49d9-9af4-d75b1ac27d32" />

### Step 10 — Fixed blank pages
The report body was 9.42in wide, wider than the 8.5in printable page — this was silently causing blank pages after every content page. Fixed by resizing the body/table width to fit within the page.

<img width="1440" height="852" alt="image" src="https://github.com/user-attachments/assets/c5856759-8252-4713-a0cf-9cde3cc338c9" />


### Step 11 — Interactivity
- **Expand/collapse per country:** clicking a country name toggles its monthly detail rows open/closed. Collapsed by default when the report first loads.
- **Click-to-sort:** clicking the "total new confirmed" column header sorts all countries by their total case count, using:
  ```
  =Sum(Fields!total_new_confirmed.Value)
  ```
<img width="1440" height="856" alt="image" src="https://github.com/user-attachments/assets/9e829eab-bf92-4c7b-b493-d36112a2cdd4" />

<img width="1440" height="854" alt="image" src="https://github.com/user-attachments/assets/fa15968a-9936-43c1-bed3-d5a3d9ae3400" />

### Step 12 — Publish
Published to the `covid19_project_workspace` Fabric workspace, where it appears as a **Paginated Report** item. Verified in the Power BI Service that everything (parameters, sorting, expand/collapse, data bars, colors, and repeating headers) still worked correctly after publishing.

<img width="1440" height="804" alt="image" src="https://github.com/user-attachments/assets/f7770188-7e0d-4671-bb60-ebb0592dccef" />

<img width="1440" height="814" alt="image" src="https://github.com/user-attachments/assets/6a3eabf8-4493-40e4-a66e-9c38d4227527" />

---

## Key Things I Learned

- **MDX parameters aren't like SQL parameters** — if a filter uses `STRTOSET(..., CONSTRAINED)`, the parameter must pass the exact MDX unique name, not a plain value. The fix is a helper dataset that fetches both the friendly label and the technical unique name.
- **`RepeatOnNewPage` only works on Row Group static members** — the same-looking setting exists on Column Groups too, but using it there breaks the report.
- **Data bars and indicators scale per-row by default**, which makes every bar look "full" unless you manually set a Minimum/Maximum — and the right scope (whole dataset vs. per-group) depends on what you're trying to compare.
- **Paginated reports need non-interactive login credentials** once published — the interactive sign-in used while building the report locally doesn't work when the report runs unattended in the cloud.
- **A report body wider than the page causes blank pages**, not an error — worth checking early if pagination looks broken after adding new columns or fields.

## Tech Stack
Power BI Report Builder (.rdl), Microsoft Fabric (Semantic Models, XMLA endpoint), Azure SQL Database, Azure Analysis Services connection type, MDX, DAX/Direct Lake (underlying semantic model), Power BI Service.
