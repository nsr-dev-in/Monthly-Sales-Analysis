# 📊 Monthly Sales Analysis Dashboard using Excel

## 📌 Project Overview

This project presents a comprehensive **Monthly Sales Analysis Dashboard** built using **Microsoft Excel**. The dashboard analyzes sales performance across 12 months, tracks Month-over-Month (MoM) growth, highlights key business metrics, and provides valuable insights through interactive visualizations.

The primary objective of this project is to demonstrate how Excel can be used as a powerful data analytics tool for business decision-making.

---

## 🎯 Project Objectives

* Analyze monthly sales performance.
* Calculate Month-over-Month (MoM) Growth.
* Identify top and bottom-performing months.
* Generate key business insights.
* Visualize sales trends using charts and dashboards.

---

## 🛠️ Tools & Technologies Used

* **Microsoft Excel**
* **Excel Functions**
* **Conditional Formatting**
* **Data Visualization Charts**

---

## 📂 Dataset Information

The dataset contains monthly sales data for a single year.

| Column Name    | Description                                      |
| -------------- | ------------------------------------------------ |
| Month          | Name of the month                                |
| Sales          | Total sales generated during the month           |
| MoM Growth (%) | Percentage growth compared to the previous month |

---

## 📈 Key Performance Indicators (KPIs)

The dashboard includes the following business metrics:

* ✅ Total Sales
* ✅ Average Sales
* ✅ Maximum Sales
* ✅ Minimum Sales
* ✅ Best Performing Month
* ✅ Worst Performing Month
* ✅ Month-over-Month Growth (%)

---

## 🧮 Excel Functions Used

```excel
SUM()
AVERAGE()
MAX()
MIN()
XLOOKUP()
```

---

## 📐 Formula Documentation

### Month-over-Month (MoM) Growth

```excel
=(Current Month Sales - Previous Month Sales)/Previous Month Sales
```

Example:

```excel
=(B3-B2)/B2
```

### Total Sales

```excel
=SUM(B2:B13)
```

### Average Sales

```excel
=AVERAGE(B2:B13)
```

### Maximum Sales

```excel
=MAX(B2:B13)
```

### Minimum Sales

```excel
=MIN(B2:B13)
```

### Best Performing Month

```excel
=XLOOKUP(MAX(B2:B13),B2:B13,A2:A13)
```

### Worst Performing Month

```excel
=XLOOKUP(MIN(B2:B13),B2:B13,A2:A13)
```

---

## 📊 Dashboard Features

* Monthly Sales Trend Analysis
* Month-over-Month Growth Tracking
* KPI Summary Metrics
* Conditional Formatting for Growth Analysis
* Line Chart for Sales Trend
* Bar Chart for Monthly Comparison
* Business Insights Generation

---

## 🔍 Key Insights

* Annual sales reached **₹1,478,000**.
* **December** recorded the highest sales (**₹180,000**).
* **January** recorded the lowest sales (**₹85,000**).
* Negative growth was observed in **July (-5.60%)** and **September (-3.03%)**.
* Overall sales showed a strong upward trend throughout the year.

---

## 📁 Project Structure

```text
monthly-sales-analysis-dashboard/
│
├── data/
│   └── monthly_sales_data.xlsx
│
├── dashboard/
│   └── sales_dashboard.xlsx
│
├── screenshots/
│   └── dashboard_preview.png
│
├── README.md
├── LICENSE
└── .gitignore
```

---

## 🚀 How to Use

1. Clone this repository.

```bash
git clone https://github.com/yourusername/monthly-sales-analysis-dashboard.git
```

2. Open the Excel workbook.
3. Explore the dashboard and analyze sales trends.
4. Modify the dataset to perform your own analysis.

---

## 🎓 Learning Outcomes

Through this project, I learned:

* Business Metrics Analysis
* KPI Creation in Excel
* Data Cleaning & Preparation
* Dashboard Design Principles
* Data Visualization Techniques
* Business Insight Generation

---

## 🤝 Connect With Me

If you found this project useful, feel free to connect with me on LinkedIn and explore my other Data Analytics projects.

⭐ If you like this project, don't forget to star the repository!
