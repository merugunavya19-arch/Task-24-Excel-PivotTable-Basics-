# Task 24 Excel PivotTable Basics

## Project Overview

This project focuses on analyzing sales data using Excel PivotTables. The main objective is to summarize sales information by product category, region, and month to understand sales performance and identify useful business insights.

This project was completed as part of my Data Analytics Internship at **Veda Technology**.

## Objectives

- Learn the basics of Excel PivotTables.
- Summarize total sales by product category.
- Compare sales performance across different regions.
- Analyze monthly sales trends.
- Identify the highest- and lowest-performing categories and regions.
- Present key findings in a clear and simple format.

## Tools and Technologies

- Microsoft Excel
- Excel PivotTables
- Sample Superstore Dataset

## Dataset Description

The project uses the Sample Superstore sales dataset, which contains information about orders, customers, products, sales, profit, and regions.

Important columns used in this analysis include:

- `Order Date` — Date when an order was placed.
- `Region` — Geographical region associated with the order.
- `Category` — Product category.
- `Sales` — Sales amount for each order item.

The original dataset is preserved, and PivotTables are created from the source data.

## Project Implementation

### 1. Sales by Category

**Rows:** Category  
**Values:** Sum of Sales

This PivotTable summarizes total sales for each product category and helps identify which category contributes the most and least sales.

### 2. Sales by Region

**Rows:** Region  
**Values:** Sum of Sales

This PivotTable compares total sales across regions and helps identify the highest- and lowest-performing regions.

### 3. Monthly Sales Trend

**Rows:** Order Date, grouped by Months and Years  
**Values:** Sum of Sales

This PivotTable summarizes monthly sales over time. Grouping by both month and year helps distinguish the same month across different years.

## Key Insights

1. **Sales by Category:** Technology recorded the highest total sales of **$836,154.03**, while Furniture recorded the lowest total sales of **$741,999.80**.

2. **Sales by Region:** The West region recorded the highest total sales of **$725,214.66**, while the South region recorded the lowest total sales of **$391,721.91**.

3. **Monthly Sales Trend:** November 2017 recorded the highest monthly sales of **$118,447.83**.

*Note: These figures should be checked against the final workbook before submission.*

## Key Learnings

- Learned how to create PivotTables in Excel.
- Understood how to place fields in Rows and Values.
- Practiced summarizing sales using Sum of Sales.
- Learned how to group dates by month and year.
- Improved skills in comparing sales performance.
- Learned how to convert summarized data into business insights.

## Conclusion

This project demonstrates how Excel PivotTables can simplify sales data analysis. By summarizing sales by category, region, and month, it becomes easier to compare performance, identify trends, and support data-driven business decisions.

## Project Structure

```text
Excel-PivotTable-Basics/
├── README.md
├── Excel_PivotTable_Basics.xlsx
└── SampleSuperstore.xlsx
```

The source dataset filename may differ depending on the file supplied. Include it in the repository only if you are permitted to share it.

## Internship Information

**Organization:** Veda Technology  
**Task:** Excel PivotTable Basics  
**Domain:** Data Analytics  
**Author:** M. Navya Sri
