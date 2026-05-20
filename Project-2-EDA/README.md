# Ecommerce Exploratory Data Analysis

This project completes DecodeLabs Data Analytics Project 2: Exploratory Data Analysis.

The work analyzes 1,200 ecommerce order records to identify descriptive statistics, trends, outliers, product performance, referral-source performance, payment behavior, and operational improvement opportunities.

## Live Preview

Open `index.html` in a browser, or enable GitHub Pages for this repository and use the Pages link.

## Final Deliverables

- `outputs/data-analytics-p2/Data Analytics P2 - Ecommerce EDA Dashboard.xlsx`
- `outputs/data-analytics-p2/Data Analytics P2 - Investor Brief.pdf`

## What Is Included

- Executive dashboard with KPI cards and charts
- EDA statistics: count, mean, median, min, max, standard deviation, quartiles, and outlier fence
- Segment analysis by product, referral source, payment method, order status, and month
- High-value outlier table
- Raw data table with added month and quarter fields
- Methodology sheet for auditability
- Investor-style PDF summary

## Key Findings

- The dataset contains 1,200 orders worth `$1,264,761.96`.
- Average order value is about `$1,054`.
- Chair, Printer, Laptop, Tablet, and Monitor are all meaningful revenue contributors, so the business is not dependent on one product only.
- Instagram is the strongest referral source by recorded order value.
- Returned and cancelled orders represent a major improvement opportunity.
- Eight orders are high-value outliers above the IQR upper fence.

## How To Rebuild

The generated workbook and PDF can be rebuilt from the Python script:

```powershell
python build_final_package.py
```

If your local Python does not have the required packages, install:

```powershell
pip install -r requirements.txt
```

The script expects the source dataset at:

```text
C:\Users\prais\Downloads\Dataset for Data Analytics (1).xlsx
```

## Repository Notes

The final files are already included under `outputs/data-analytics-p2/`, so the project can be reviewed without rerunning the script.
