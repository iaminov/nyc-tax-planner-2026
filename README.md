# 2026 NYC Tax, Solo 401(k), HSA & ACA Premium Tax Credit Planner

A comprehensive planning and modeling tool designed for self-employed individuals (sole proprietors / single-member LLCs) residing in New York City.

The application calculates your maximum traditional **Solo 401(k)** contribution, models personal **HSA** contributions with prorated monthly eligibility and the December last-month rule, solves the circular interaction between the **Self-Employed Health Insurance Deduction** and the **ACA Premium Tax Credit (PTC)**, and generates estimated quarterly tax vouchers for **Federal**, **New York State**, **New York City**, and the **Metropolitan Commuter Transportation Mobility Tax (MCTMT)**.

Inputs can be loaded from an Excel workbook (`tax_inputs_template.xlsx`) and modified interactively in a graphical user interface (GUI) to model real-time "what-if" tax scenarios.

---

## Key Features

1. **Solo 401(k) Optimization**:
   - Computes net earnings from self-employment ($92.35\%$), deductible half of SE tax, and adjusted earned income.
   - Determines maximum employee elective deferrals ($24,500$ in 2026, plus age 50+ catch-up of $8,000$ or age 60–63 enhanced catch-up of $11,250$), accounting for deferrals made at other W-2 employers.
   - Computes employer non-elective profit-sharing contributions ($20\%$ of adjusted net earnings) subject to the IRC §415(c) $72,000 annual addition limit.
   - Allows entering an actual planned contribution to model its downstream tax impact.

2. **Circular Solver: Health Insurance Deduction & ACA PTC**:
   - IRS Form 7206 and IRC §162(l) limit the self-employed health insurance deduction to net earned business profit minus retirement contributions, as well as net out-of-pocket premiums after subsidies.
   - At the same time, the health insurance deduction lowers AGI and ACA MAGI, which in turn alters the benchmark subsidy (PTC) on Form 8962.
   - Rather than naive fixed-point iteration (which often fails or oscillates across subsidy cliffs), this planner partitions income at statutory FPL boundaries (100%, 133%, and 400% FPL) and performs piecewise bisection ($75$ iterations, precision $< 10^{-20}$) to identify the exact optimal deduction and credit.
   - Incorporates 2026 statutory rules, including the expiration of ARP/IRA enhanced subsidies (400% FPL cliff) and the elimination of excess APTC repayment caps.

3. **HSA Monthly Tracking**:
   - Prorates personal HSA allowances ($4,400 self-only / $8,750 family in 2026, plus $1,000 catch-up if 55+) based on first-of-month eligibility.
   - Supports the IRC §223(b)(8) December last-month rule and displays a detailed month-by-month schedule tab.
   - Properly accounts for employer/pretax payroll HSA contributions without double-deducting.

4. **Metropolitan Commuter Transportation Mobility Tax (MCTMT) Nuances**:
   - **Net SE Earnings Base**: Evaluated strictly on net earnings from self-employment (IRC §1402(a)), computed as $\text{Schedule C profit} \times 0.9235$.
   - **Exclusions**: W-2 wages and capital gains do **not** count toward the MCTMT threshold, and Solo 401(k) contributions do **not** reduce it.
   - **Per-Zone Evaluation**: The threshold is evaluated separately for Zone 1 (NYC: Manhattan, Brooklyn, Queens, Bronx, Staten Island) and Zone 2 (Nassau, Suffolk, Westchester, Rockland, Orange, Putnam, Dutchess).
   - **Cliff Taxation**: If net earnings in a zone exceed the threshold, the tax rate ($0.60\%$ Zone 1, $0.34\%$ Zone 2) applies to the **entire** allocated net earnings in that zone, not just the excess over the threshold.
   - **2026 vs. 2025 Law Change**: The $150,000 threshold applies beginning in tax year 2026. For tax year 2025 returns (even if prepared in 2026), the prior $50,000 threshold applies. Both years can be modeled directly.

5. **NYC Unincorporated Business Tax (UBT) Full Modeling**:
   - **Separate NYC Tax**: Administered by the NYC Department of Finance (Forms NYC-202 / NYC-202S), completely separate from regular NYC personal income tax and MCTMT.
   - **Gross Filing Threshold ($95,000)**: Evaluated on **gross business receipts**, not net profit. Businesses with $> \$95,000$ in gross revenue must generally file a return even if actual tax liability is $\$0$.
   - **Personal Services Deduction**: Automatically deducts $20\%$ of net business profit up to a **$10,000** maximum allowance for services rendered by the proprietor (NYC Admin Code § 11-508(a)).
   - **Specific Business Exemption**: Applies the statutory **$5,000** annual exemption.
   - **4% Tax & Small Business Credit**:
     - Gross UBT calculated at $4\%$.
     - **Full Elimination ($\le \$3,400$)**: For profits up to $\$100,000$, taxable income is $\$85,000$ ($\$100\text{k} - \$10\text{k} - \$5\text{k}$), gross tax is $\$3,400$, and the small-business credit eliminates $100\%$ of the tax ($-\$3,400$), resulting in **$0 UBT owed**.
     - **Sliding Phaseout**: Taxes between $\$3,400$ and $\$5,400$ receive a partial credit; taxes $\ge \$5,400$ receive no credit.
   - **Form NYC-202 vs. NYC-202S**: Alerts the user if traditional Solo 401(k) contributions or multiple trades require filing the full Form NYC-202 rather than the simplified Form NYC-202S.
   - **NYC Resident Credit (Form IT-219)**: Offsets NYC personal income tax by the allowed percentage ($100\%$ down to $15\%$) of any UBT actually paid.

6. **Quarterly Tax Estimator**:
   - Computes annual liabilities and divides remaining obligations (net of withholdings and prior estimated payments) across remaining installment dates.
   - Covers Federal income tax, SE tax, 0.9% Additional Medicare tax (Form 8959), 3.8% Net Investment Income Tax (NIIT), NYS income tax, NYC resident personal tax (with school credits and UBT credit), MCTMT, and NYC UBT.

---

## Repository Structure

```
├── nyc_tax_planner_2026.py      # Main application script & Tkinter GUI
├── tax_inputs_template.xlsx     # Clean Excel template for user inputs
├── requirements.txt             # Python dependencies (openpyxl)
├── .gitignore                   # Ignores local user workbooks & personal data
└── README.md                    # Documentation and usage guide
```

---

## Getting Started

### Prerequisites

- Python 3.10 or higher
- `openpyxl` (installed via requirements)

### Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/iaminov/nyc-tax-planner-2026.git
   cd nyc-tax-planner-2026
   ```

2. Create and activate a virtual environment:
   ```bash
   # Windows (PowerShell)
   python -m venv .venv
   .venv\Scripts\Activate.ps1

   # macOS / Linux
   python3 -m venv .venv
   source .venv/bin/activate
   ```

3. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

---

## How to Fill Out the Excel Input Sheet

The repository includes a clean, pre-formatted template: **`tax_inputs_template.xlsx`**.

### Setup
1. Copy `tax_inputs_template.xlsx` to `tax_inputs_2026.xlsx`:
   ```powershell
   # Windows
   Copy-Item tax_inputs_template.xlsx tax_inputs_2026.xlsx

   # macOS / Linux
   cp tax_inputs_template.xlsx tax_inputs_2026.xlsx
   ```
   *(Note: If `tax_inputs_2026.xlsx` does not exist when you launch the script, the application will automatically create it from `tax_inputs_template.xlsx`.)*

2. Open `tax_inputs_2026.xlsx` in Microsoft Excel, LibreOffice Calc, or Google Sheets.

3. Edit only the yellow **Value (Enter Here)** column (Column D).

### Sections Overview

| Section | Key Fields to Fill Out | Notes |
| :--- | :--- | :--- |
| **Business & W-2** | `gross_business`, `business_expenses`, `w2_box1`, `w2_social_security_wages`, `w2_medicare_wages`, `w2_prior_401k_deferrals`, `short_gains`, `long_gains`, `qualified_dividends`, `interest_other_ordinary` | Enter annual gross Schedule C revenue and ordinary deductible expenses. Box 3 and Box 5 wages prevent overcounting Social Security SE tax and Additional Medicare. |
| **Retirement & HSA** | `age`, `solo_actual`, `hsa_actual`, `employer_hsa`, `hsa_first_month`, `hsa_last_month`, `hsa_family`, `use_hsa_last_month_rule` | Set your planned traditional Solo 401(k) contribution and personal HSA contribution. The planner verifies contributions against allowable caps. |
| **Marketplace / Insurance** | `enrollment_monthly`, `benchmark_monthly`, `aptc_monthly`, `aca_first_month`, `aca_last_month`, `self_emp_coverage_eligible`, `qhp_ptc_eligible` | `benchmark_monthly` defaults to $868.20 (SLCSP benchmark). `enrollment_monthly` is your actual plan premium. Enter any APTC monthly credit received. |
| **New York & Other** | `mct_zone1_percent`, `mct_zone2_percent`, `tax_year`, `mct_threshold_override`, `ubt_auto_estimate`, `qbi_enabled`, `federal_standard_override`, `ny_standard_override` | Set `mct_zone1_percent` to 100 for NYC/Brooklyn business activity. `tax_year` controls the MCTMT threshold ($150k for 2026 vs $50k for 2025). |
| **Withholdings & Already Paid**| `fed_withholding`, `nys_withholding`, `nyc_withholding`, `federal_estimates_already_paid`, `ny_estimates_already_paid`, `installments_remaining` | Enter W-2 withholdings and quarterly payments already remitted to project remaining installment cash requirements. |

> **Formulas Supported**: You can write standard Excel formulas in Column D (e.g., `=125000-15000` or `=SUM(B10:B12)`). The planner reads evaluated values directly.

---

## Running the Application

### 1. Launch the Planner GUI
```bash
# Uses default tax_inputs_2026.xlsx
python nyc_tax_planner_2026.py
```

Or target a specific custom file:
```bash
python nyc_tax_planner_2026.py --excel path/to/my_inputs.xlsx
```

### 2. Live Session Overrides (Modeling Scenarios)
- When the GUI loads, all fields are populated from your Excel workbook.
- **You can freely edit any value directly in the GUI** to test scenarios (e.g., "What if my profit increases by $25k?" or "What if I maximize my Solo 401(k)?").
- Click **Calculate** to recompute all taxes and credits immediately.
- The **Results tab** displays an active session diff comparing your modified GUI values against your original Excel values.
- **Controls**:
  - **`Calculate`**: Evaluates using current GUI inputs.
  - **`Reload Excel`**: Reverts all GUI fields back to the saved Excel file.
  - **`Load Excel…`**: Select and switch to any other Excel file.
  - **`Save to Excel…`**: Persists your GUI session modifications back into an Excel file.

### 3. Running Automated Tests
The repository includes a comprehensive unit test suite covering math calculations, bisection convergence, MCTMT cliff rules, and Excel round-trip parsing:
```bash
python nyc_tax_planner_2026.py --test
```

---

## Privacy & Personal Data Security

- **No Personal Data Committed**: The repository only tracks `tax_inputs_template.xlsx` containing clean generic defaults.
- **Git Ignored**: `.gitignore` is configured to ignore active user workbooks (`tax_inputs_2026.xlsx`, `my_*.xlsx`, `*.json`) to prevent financial data from being pushed to version control.

---

## Disclaimer

This software is an educational planning model and cash-flow estimation tool. It is not formal tax advice, legal counsel, or certified return preparation software. Tax laws, phaseouts, and state-specific regulations are subject to administrative updates. Always review your final return with a qualified CPA or tax professional before filing.
