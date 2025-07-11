# Project Overview

This project contains two Excel files 

- **File 1: Annuity Calculator** — an interactive VBA-powered tool for solving Time Value of Money (TVM) problems involving annuities.
- **File 2: Insurance Charges Visualization** — a dynamic Excel dashboard analyzing insurance charges across demographic and regional factors.





# File 1: Annuity Calculator 

## Workbook Overview

The Annuity Calculator is an interactive Excel-based tool built using VBA UserForms. It
enables users to solve Time Value of Money (TVM) problems involving annuities. Users can
calculate missing financial variables, including Future Value (FV), Present Value (PV),
Payment (PMT), Number of Years (n), or Annual Interest Rate (i) based on selected rate
types and payment frequencies.

## Key Features

```
● User-Friendly Interface: Clean, form-based input for ease of use.
● Solves for Any Variable: Calculate FV, PV, PMT, i, or n when the others are known.
● Supports Multiple Rate Types:
    ○ Nominal or Effective annual rates.
    ○ Adjustable payment frequencies (Annual, Semiannual, Quarterly, Monthly).
● Ordinary and Due Annuity Support:
    ○ Choose whether payments occur at the start or end of each period.
● Interest Rate Solver: Uses Newton-Raphson iteration to calculate unknown interest
rates.
```


# File 2: Insurace Charges Visualization 

## Workbook Overview

This Excel workbook aims to showcase my ability to analyze and visualize data using core Excel techniques. The project demonstrates skills in:

```
● Data cleaning to structure, categorize, and prepare raw data for analysis

● Summary metrics (averages, VLOOKUP, IF-based calculations)

● Conditional formatting to highlight key data trends

● Interactive dashboards using PivotTables, slicers, and charts to explore
demographic and regional factors influencing insurance charges.
```

## Worksheet Overview

Below is an overview of each worksheet in the order they appear in the workbook:

### Project Overview

```
● Contains author information, date, and a summary of the project’s purpose.

● Includes a clickable link to the dataset source.

● Provides a Table of Contents with hyperlinks and descriptions of each worksheet.
```

### Original Dataset

```
● Displays the raw data formatted as an Excel Table for easy analysis and referencing
in formulas and PivotTables.
```

### Conditional Formatting

```
Applies customized conditional formatting to make key trends visually clear:

● BMI: In-cell horizontal bar charts.

● Children:

    ○ 0 children → red circle 
   
    ○ 1–2 children → yellow circle 

    ○ 3+ children → green circle

● Smoker: Cells containing "yes" are highlighted in red 

● Region: Region names are colored for easy distinction:

    ○ Southwest → Brown

    ○ Southeast → Peach/Tan

    ○ Northwest → Dark Purple

    ○ Northeast → Light Purple

● Charges: Conditional color scale:

    ○ Low charges → Green

    ○ Medium charges → Yellow/Orange

    ○ High charges → Red
```

### Dashboard

```
● An Interactive Dashboard with various charts, slicers, and graphics allowing users to
explore trends through a clean, engaging interface
```

### Age and Smoker Status

```
● Calculates the average insurance charge grouped by Age group and Smoker status

● Displays results using a clustered bar chart.

● Includes slicers for dynamic filtering.
```

### Charges by Region

```
● Calculates the average insurance charge by region.

● Displays results using a bar graph.

● Includes slicers for dynamic filtering.
```

### Count by Region

```
● Calculates the count of individuals per region.

● Displays results using a pie chart.

● Includes slicers for dynamic filtering.
```

