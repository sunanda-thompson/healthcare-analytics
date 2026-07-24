# Medicare Provider Utilization & Payment Analysis, CY2024

**Healthcare Reporting Analyst Portfolio Project**
Data Source: [CMS Medicare Physician & Other Practitioners by Provider and Service](https://data.cms.gov/provider-summary-by-type-of-service/medicare-physician-other-practitioners/medicare-physician-other-practitioners-by-provider-and-service), Calendar Year 2024

---

## 📌 Project Overview

This project analyzes **352,948 Medicare fee-for-service (FFS) records** published by the Centers for Medicare & Medicaid Services (CMS) for CY2024, covering four high-volume specialties: **Internal Medicine, Family Practice, Cardiology, and Psychiatry**, across three major Southern markets: **Texas, Florida, and Georgia**.

These states and specialties represent core markets for large national payers (UnitedHealth Group, Elevance Health, Humana, CVS Health/Aetna, Molina Healthcare), and the analysis mirrors the descriptive reporting, cost/payment analysis, geographic comparison, and provider benchmarking work performed by Healthcare Reporting Analysts on payer network management teams.

**Tools used:** Excel (data cleaning, pivot tables, exploratory analysis) → Power BI (5-page interactive operational scorecard)

---

## 🎯 Hypothesis

> Medicare provider utilization and payment data for 2024 reveals meaningful differences in service volume, cost, and provider performance across specialties and geographies. By analyzing CMS fee-for-service claims data across Internal Medicine, Family Practice, Cardiology, and Psychiatry in Texas, Florida, and Georgia, a Healthcare Reporting Analyst can identify which specialties drive the highest costs, which states show the greatest utilization, and where provider payment gaps exist: insights directly relevant to payer network management and operational reporting.

## ❓ Guiding Business Questions

1. Which specialty has the highest total services and beneficiaries served across TX, FL, and GA?
2. What is the average Medicare allowed amount and payment by specialty and state, and which specialty has the largest gap between submitted charges and Medicare payment?
3. Using standardized payment amounts, which state has the highest cost per service when geographic differences are removed?
4. Which providers are top performers by total beneficiaries served within their specialty, and how does Medicare participation vary by specialty and state?

---

## 🖥️ Dashboard Preview
> Full interactive dashboard: [`Medicare PowerBI Dashboard.pdf`](Medicare%20PowerBI%20Dashboard.pdf). Individual finding charts below live in the `Charts/` folder.

### Finding 1: Specialty Utilization
Internal Medicine leads total services and beneficiaries served in every state.
![Total Services by Specialty and State](https://raw.githubusercontent.com/sunanda-thompson/healthcare-analytics/65c1f8187df74df22b69a71b1a7b5f71137198aa/Medicare-Provider-Service-2024/Charts/1.1%20Total%20Services.png)
![Total Beneficiaries Served by Specialty and State](https://raw.githubusercontent.com/sunanda-thompson/healthcare-analytics/65c1f8187df74df22b69a71b1a7b5f71137198aa/Medicare-Provider-Service-2024/Charts/1.2%20Total%20Beneficiaries.png)

### Finding 2: Cost & Payment Gap
Cardiology shows the widest gap between submitted charges and Medicare payment (5.67x ratio).
![Average Submitted Charge vs Medicare Allowed vs Medicare Payment](https://raw.githubusercontent.com/sunanda-thompson/healthcare-analytics/65c1f8187df74df22b69a71b1a7b5f71137198aa/Medicare-Provider-Service-2024/Charts/2.%20Average%20Amount.png)

### Finding 3: Geographic Cost Comparison
Florida has the highest standardized payment even after geographic adjustment, and Cardiology commands the highest payment in every state.
![Average Standardized Medicare Payment by State and Specialty](https://raw.githubusercontent.com/sunanda-thompson/healthcare-analytics/65c1f8187df74df22b69a71b1a7b5f71137198aa/Medicare-Provider-Service-2024/Charts/3.%20Average%20Medicare%20Payments.png)

### Finding 4: Provider Performance
Eric Weiner (Internal Medicine, FL) tops the list with 73,417 beneficiaries served.
![Top 20 Providers by Total Beneficiaries Served](https://raw.githubusercontent.com/sunanda-thompson/healthcare-analytics/65c1f8187df74df22b69a71b1a7b5f71137198aa/Medicare-Provider-Service-2024/Charts/4.%20Top%2020%20Providers.png)

---

## 🔑 Key Findings

**Scope:** 64,020,716 total services · 28,611,681 total beneficiaries served · $71.71 average standardized Medicare payment · 5.15 average charge-to-payment ratio across TX, FL, and GA (2024)

1. **Internal Medicine leads utilization in every state.** 31.1M total services and 12.0M beneficiaries served across TX, FL, and GA, more than any other specialty. Family Practice is a distant second (19.8M services), followed by Cardiology (11.3M) and Psychiatry (1.8M).

2. **Florida dominates utilization overall.** Florida accounts for the largest share of total services across all four specialties, ahead of Texas and Georgia, driven primarily by its large Internal Medicine and Family Practice volumes.

3. **Cardiology has by far the widest gap between what's charged and what's paid.** Average submitted charge is $425 vs. an average Medicare payment of $107, a **5.67x charge-to-payment ratio**, the highest of any specialty (vs. 5.4x for Psychiatry, 4.8x for Internal Medicine, 3.3x for Family Practice).

4. **Florida has the highest cost per service even after geographic adjustment.** Using standardized payment amounts (which remove regional cost-of-living differences), Florida still shows the highest average standardized payment of the three states, meaning the cost difference isn't just a geography artifact.

5. **Cardiology commands the highest standardized payment in every single state** (FL, GA, and TX), confirming that Cardiology's higher cost profile holds regardless of location.

6. **Medicare participation is nearly universal.** 99.98% of providers in this dataset are Medicare-participating (352,889 of 352,948 records), with only 0.02% non-participating.

7. **Top individual provider by beneficiaries served:** Eric Weiner (Internal Medicine, FL), 73,417 beneficiaries and 101,862 total services, the highest of any provider in the dataset. Emmanuel Cruz Caban (also FL Internal Medicine) is a close second at 70,227 beneficiaries.

---

## 🧹 Data Cleaning Summary

Full documentation lives in the workbook's **Data Handling Summary** tab. At a high level:
- Raw CMS export reduced from 28 columns down to **21 analysis-ready columns** (Provider ID, name, location, specialty, participation status, procedure code/description, place of service, utilization counts, and charge/payment fields).
- Filtered to Internal Medicine, Family Practice, Cardiology, and Psychiatry providers in TX, FL, and GA.
- **352,948 clean records** retained; original unmodified source data preserved in a separate **Raw Data** tab for reference/audit.

## 📖 Data Dictionary

See the **Data Dictionary** tab in the workbook for full column-level definitions (name, data type, description) across all 21 columns in the clean dataset.

---

## 📁 Repository Structure

```
Medicare-Provider-Service-2024/
├── README.md
├── Medicare_Provider_and_Service_2024_Working_Copy.xlsx   # Excel workbook: cleaning, pivots, findings
├── Medicare PowerBI Dashboard.pdf                          # Power BI dashboard export
└── Charts/
    ├── 1_1_Total_Services.png
    ├── 1_2_Total_Beneficiaries.png
    ├── 2__Average_Amount.png
    ├── 3__Average_Medicare_Payments.png
    └── 4__Top_20_Providers.png
```

## 🛠️ Tools & Skills Demonstrated

- **Excel:** Data cleaning, PivotTables, exploratory data analysis, chart-building
- **Power BI:** Data modeling, DAX-free descriptive dashboards, multi-page interactive report design, filter/slicer design
- **Analytical skills:** Descriptive reporting, cost/payment gap analysis, geographic benchmarking, provider performance scorecards, directly aligned to payer Healthcare Reporting Analyst workflows
