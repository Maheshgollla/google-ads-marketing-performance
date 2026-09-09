# Google Ads Marketing Performance Analytics

## 1. Business Problem

Analyze Google Ads campaign performance to identify high-performing campaigns, understand funnel efficiency, measure advertising ROI, and provide data-driven recommendations for campaign optimization and budget allocation.

## 2. Project Objectives

The project will evaluate Google Ads marketing performance across campaigns and key analytical dimensions to understand:

* Campaign profitability and return on advertising spend
* Conversion and revenue generation
* Advertising funnel efficiency
* Performance differences across devices, locations, and keywords
* Performance trends over time
* Opportunities for campaign scaling, optimization, and budget reduction

## 3. Key Business Questions

1. Which campaigns generate the best ROAS?
2. Which campaigns generate the most conversions and revenue?
3. Which devices perform best?
4. Which keywords drive efficient conversions?
5. Where are the largest funnel drop-offs?
6. Which campaigns should be scaled, optimized, or reduced?
7. How does performance change over time?

## 4. Dataset

### Source Dataset

The project uses the following raw dataset:

```text
data/raw/GoogleAds_DataAnalytics_Sales_Uncleaned.csv
```

The raw dataset will be preserved as the source dataset and will not be modified during the analytical workflow.

### Dataset Fields

The dataset contains advertising, funnel, revenue, campaign, geographic, device, keyword, and date information.

Detailed data profiling and quality assessment will be performed during Phase 2.

## 5. Data Structure

### Analytical Grain

One row represents an individual advertising record identified by `Ad_ID`, with associated campaign, date, location, device, keyword, advertising metrics, funnel metrics, and sales amount.

The dataset contains the following fields:

| Category             | Fields                                    |
| -------------------- | ----------------------------------------- |
| Record Identifier    | `Ad_ID`                                   |
| Campaign Dimension   | `Campaign_Name`                           |
| Date Dimension       | `Ad_Date`                                 |
| Geographic Dimension | `Location`                                |
| Device Dimension     | `Device`                                  |
| Keyword Dimension    | `Keyword`                                 |
| Advertising Metrics  | `Clicks`, `Impressions`, `Cost`           |
| Funnel Metrics       | `Leads`, `Conversions`, `Conversion Rate` |
| Revenue Metric       | `Sale_Amount`                             |

### Why the Grain Matters

The analytical grain defines what a single observation represents in the dataset. All subsequent aggregations, KPI calculations, validation procedures, and dimensional analyses must respect this grain to avoid double-counting or incorrect metric calculations.

`Ad_ID` is treated as the record-level identifier for the advertising observation. The uniqueness and integrity of `Ad_ID` will be formally validated during the Data Profiling & Audit phase.

## 6. KPIs & Metric Definitions

The following KPIs are the official metric definitions for this project.

| KPI                    | Formula                     | Unit / Format           |
| ---------------------- | --------------------------- | ----------------------- |
| CTR                    | Clicks / Impressions        | Percentage              |
| CPC                    | Cost / Clicks               | Currency per click      |
| CPL                    | Cost / Leads                | Currency per lead       |
| Conversion Rate        | Conversions / Clicks        | Percentage              |
| CPA                    | Cost / Conversions          | Currency per conversion |
| ROAS                   | Sale_Amount / Cost          | Multiple, e.g. 4.5x     |
| ROI                    | (Sale_Amount - Cost) / Cost | Percentage              |
| Revenue / Conversion   | Sale_Amount / Conversions   | Currency per conversion |
| Lead → Conversion Rate | Conversions / Leads         | Percentage              |

### Metric Calculation Rules

* Rates should generally be calculated from aggregated numerators and denominators rather than by averaging row-level rates.
* Zero denominators must return `0` or `NULL/NaN` according to the metric's business definition.
* `Sale_Amount` is treated as sales/revenue generated from the advertising activity.
* `Cost` is treated as advertising spend.
* ROAS is expressed as a multiple. For example, a ROAS of `4.5x` means that every 1 unit of advertising spend generated 4.5 units of sales.
* ROI is expressed as a percentage. For example, an ROI of `350%` means the return above advertising cost is 3.5 times the advertising cost.
* KPI values will be calculated from the dataset only in later analytical phases.
* No expected KPI values are hard-coded during Phase 1.

### Aggregation Principle

For aggregated analysis, metrics with numerators and denominators should be calculated using the aggregated underlying values.

For example:

```text
Aggregated CTR = Total Clicks / Total Impressions

Aggregated Conversion Rate = Total Conversions / Total Clicks

Aggregated ROAS = Total Sale_Amount / Total Cost
```

This prevents misleading results that can occur when averaging individual row-level rates.

## 7. Analytical Dimensions

The project will analyze Google Ads performance across the following primary dimensions:

| Dimension | Source Column   |
| --------- | --------------- |
| Campaign  | `Campaign_Name` |
| Device    | `Device`        |
| Location  | `Location`      |
| Keyword   | `Keyword`       |
| Date      | `Ad_Date`       |

### Primary Analysis Hierarchy

The analysis will follow this hierarchy:

```text
Overall
   ↓
Campaign
   ↓
Device / Location / Keyword
   ↓
Date / Trend
```

### Dimension Definitions

* **Campaign:** Used to compare advertising performance across campaigns.
* **Device:** Used to evaluate performance across desktop, mobile, tablet, and other device categories present in the dataset.
* **Location:** Used to compare advertising performance across geographic locations.
* **Keyword:** Used to evaluate keyword-level traffic, conversion efficiency, and revenue performance.
* **Date:** Used to analyze performance trends and changes over time.

These dimensions will be used for subsequent aggregation, segmentation, trend analysis, and business-performance evaluation.

Data-quality issues such as inconsistent capitalization, spelling variations, and date-format inconsistencies will be assessed during the Data Profiling & Audit phase and will not be modified during Phase 1.

## 8. Project Workflow

The project will follow the implementation workflow documented in `project_plan.md`:

```text
Raw Data
   ↓
Profiling
   ↓
Cleaning
   ↓
Validation
   ↓
Feature Engineering
   ↓
EDA
   ↓
Campaign / Dimension Analysis
   ↓
Advanced Analysis
   ↓
Business Insights
   ↓
Power BI Dashboard
   ↓
Testing & QA
   ↓
Documentation
```

## 9. Technology Stack

### Python

```text
Python
├── Pandas
├── NumPy
├── Matplotlib
├── Seaborn
└── SciPy
```

### SQL

```text
SQL
├── Data transformation
├── Aggregation
├── Validation
└── Analytical queries
```

### Excel

```text
Excel
├── Pivot tables
├── KPI analysis
└── Supporting analysis
```

### Power BI

```text
Power BI
├── Data modeling
├── DAX
├── KPI dashboard
└── Interactive reporting
```

### Git / GitHub

```text
Git/GitHub
└── Version control and portfolio documentation
```

## 10. Repository Structure

```text
google-ads-marketing-performance/
│
├── dashboard/
│   └── screenshots/
├── data/
│   ├── raw/
│   └── processed/
├── excel/
├── notebooks/
├── reports/
│   ├── analysis/
│   └── profiling/
├── sql/
├── src/
├── tests/
│
├── .gitignore
├── LICENSE
├── project_plan.md
├── README.md
└── requirements.txt
```

The raw dataset is stored under:

```text
data/raw/
```

Generated and processed data will be handled in later project phases.

## 11. Planned Deliverables

| Phase | Primary Output                      |
| ----- | ----------------------------------- |
| 1     | Repository + README + environment   |
| 2     | Profiling report + issue log        |
| 3     | Cleaned dataset + cleaning pipeline |
| 4     | Validation report + tests           |
| 5     | Feature-engineered dataset          |
| 6     | EDA notebook                        |
| 7     | Campaign/dimension analysis         |
| 8     | Statistical analysis                |
| 9     | Business recommendations            |
| 10    | Power BI dashboard                  |
| 11    | QA/reconciliation                   |
| 12    | Final portfolio repository          |

## 12. Dashboard

The Power BI dashboard will be developed in a later project phase.

**Status:** Not yet developed.

Dashboard metrics, visualizations, filters, and design will be based on validated analytical results from the preceding phases.

## 13. Key Findings

**Status:** Not yet available.

Findings will be documented only after data profiling, cleaning, validation, EDA, and analytical phases have been completed.

## 14. Business Recommendations

**Status:** Not yet available.

Business recommendations will be developed from validated analytical findings and will not be predetermined during Phase 1.

## 15. Testing & Validation

Testing and validation will be performed during later project phases.

Planned validation activities include:

* Data-quality validation
* Record and identifier validation
* KPI calculation validation
* Aggregation and reconciliation checks
* Transformation testing
* Analytical result verification
* Power BI metric reconciliation
* Final dashboard QA

No analytical test results are reported during Phase 1.

## 16. How to Run the Project

### Prerequisites

The project requires Python and Git.

### Python Environment

Create the virtual environment:

```powershell
python -m venv .venv
```

Activate the environment on Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Install project dependencies:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### Jupyter Kernel

The project Jupyter kernel is registered as:

```text
Google Ads Marketing Analytics
```

### Project Workflow

The raw dataset is located at:

```text
data/raw/GoogleAds_DataAnalytics_Sales_Uncleaned.csv
```

Subsequent project phases will define the exact notebook execution sequence and analytical workflow.

## 17. License

This project is distributed under the license specified in the repository's `LICENSE` file.
