# Data Profiling Report

## 1. Dataset Overview

The raw dataset contains **2,600 rows** and **13 columns**. The source file was verified at `C:\Users\mahes\Desktop\project1\google-ads-marketing-performance\data\raw\GoogleAds_DataAnalytics_Sales_Uncleaned.csv` and was not overwritten.

## 2. Structure & Data Types

The raw column inventory contains: Ad_ID, Campaign_Name, Clicks, Impressions, Cost, Leads, Conversions, Conversion Rate, Sale_Amount, Ad_Date, Location, Device, Keyword. `Conversion Rate` remains unchanged as the raw column name. Numeric counts are stored as numeric types; `Cost`, `Sale_Amount`, and `Ad_Date` are raw object fields requiring later treatment.

## 3. Missing Values

The dataset contains **1,150 missing cells**. The largest gaps are `Conversion Rate` (626, 24.08%), `Sale_Amount` (139, 5.35%), and `Clicks` (112, 4.31%).

## 4. Duplicates

Complete duplicate rows: **0**. Duplicate `Ad_ID` records: **0**.

## 5. Categorical Inconsistencies

`Campaign_Name` has 4 raw values, `Location` has 4, `Device` has 9, and `Keyword` has 6. Location and device case/spelling variants can fragment grouping; keywords were identified without merging.

## 6. Date Issues

All **2,600** raw dates parse under explicit format rules and all **2,600** fall within 2024-11-01 through 2024-11-30. **1,707** use non-YYYY-MM-DD representations, including **863** day-first values requiring convention confirmation. Unparseable: **0**.

## 7. Currency/Numeric Issues

`Cost` has 97 missing records and `Sale_Amount` has 139; all non-null values in both fields are temporarily parseable. Count metrics contain no unexpected decimals, negatives, or non-parseable non-missing values. The raw conversion rate ranges from 0.015 to 0.123, consistent with decimal-proportion storage.

## 8. Business-Rule Violations

The audit evaluated rules with missing values excluded. Constraint violations were not found. The three anomaly checks were retained as source-definition checks rather than treating ordinary non-anomalous records as violations.

| Rule                  | Rule_Type   |   Records_Evaluated |   Records_Violating |   Records_Flagged |
|:----------------------|:------------|--------------------:|--------------------:|------------------:|
| Clicks <= Impressions | constraint  |                2437 |                   0 |                 0 |
| Leads <= Clicks       | constraint  |                2443 |                   0 |                 0 |
| Conversions <= Clicks | constraint  |                2417 |                   0 |                 0 |
| Cost >= 0             | constraint  |                2503 |                   0 |                 0 |
| Sale_Amount >= 0      | constraint  |                2461 |                   0 |                 0 |
| Conversion_Rate >= 0  | constraint  |                1974 |                   0 |                 0 |
| Conversions > Leads   | anomaly     |                2481 |                   0 |                 0 |
| Conversions > Clicks  | anomaly     |                2417 |                   0 |                 0 |
| Clicks > Impressions  | anomaly     |                2437 |                   0 |                 0 |

## 9. Outliers

IQR detection found **0** outliers across the six requested metrics. No outliers were removed or modified.

| Column      |   Total_Outliers |   Outlier_Percentage |
|:------------|-----------------:|---------------------:|
| Impressions |                0 |                    0 |
| Clicks      |                0 |                    0 |
| Cost        |                0 |                    0 |
| Leads       |                0 |                    0 |
| Conversions |                0 |                    0 |
| Sale_Amount |                0 |                    0 |

## 10. Data-Quality Issue Summary

The issue log contains **16** entries by severity:

| Severity   |   Issues |
|:-----------|---------:|
| High       |        3 |
| Low        |        2 |
| Medium     |       11 |

## 11. Business Impact

Missing conversion and revenue fields can affect funnel and revenue reconciliation. Date representation differences can affect time grouping. Campaign, location, and device variants can fragment dimension-level reporting. Conversion-rate inconsistencies require source-definition confirmation before recomputation.

## 12. Recommended Cleaning Actions

Apply only approved Phase 3 mappings and missing-value rules. Standardize date representation after confirming day-first interpretation. Review categorical mappings without merging keywords automatically. Recompute conversion rate only under the project-approved Phase 5 rule.

## 13. Phase 2 Conclusion

Phase 2 audit work is complete. The raw CSV and raw DataFrame were preserved; all transformations in this notebook were temporary audit objects.
