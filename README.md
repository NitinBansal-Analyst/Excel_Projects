# Excel_Project - (Automation & Visualization)
# This Excel project is designed to analyze sales and customer service data to provide valuable business insights. It consists of seven sheets, each serving a unique purpose in data analysis and visualization. Below is a brief description of each sheet:

# Visualization:
This sheet features a comprehensive dashboard created using various interactive elements like option buttons, checkboxes, combo boxes, and slicers. These controls allow users to filter and interact with the data dynamically. The dashboard showcases different types of charts, including stacked bar charts, stacked column charts, area charts, donut charts, line charts, and clustered bar charts, offering a multi-dimensional view of sales and category performance across different regions and time frames.

# Macros:
This sheet focuses on data filtering and formatting using macros. It includes dependent dropdown lists for filtering data based on selected column values and subcategories. Form control buttons with macros are used for specific actions:

**Apply Filter**: Filters data based on selected category and subcategory.
**Reset Filter**: Clears applied filters and resets the view.
**Apply Formatting**: Formats filtered or unfiltered data to enhance readability.
**Apply Date Format**: Converts date formats to mm-dd-yyyy.

# Analysis1:
This sheet uses Power Pivot to create a pivot table that combines data from the "SalesData" and "CallsHandled" sheets. Conditional formatting is applied to highlight key metrics, using symbols and data bars for visual emphasis. This sheet provides insights into sales performance and customer service efficiency.

# Analysis2:
This sheet features pivot tables based on raw data from the "SuperStore" sheet. It performs various calculations to generate insights that contribute to the visualizations in the dashboard. This helps in identifying sales trends, customer preferences, and regional performance.

# SuperStore:
Contains raw data related to sales transactions and product details. This data serves as the source for the "Analysis2" pivot tables and is crucial for deriving insights used in the dashboard.

# SalesData:
This sheet holds raw sales data, which is used in conjunction with other data sources to perform in-depth analysis and generate pivot tables in "Analysis1."

# CallsHandled:
This sheet includes raw data on customer service interactions. It is combined with sales data in "Analysis1" to evaluate the impact of customer service on overall sales performance.

# Summary
Overall, this project integrates multiple Excel functionalities, such as dashboards, macros, formulas, pivot tables, and conditional formatting, to provide a robust platform for analyzing and visualizing sales data. By effectively using Excel's advanced features, the project delivers actionable insights that support data-driven decision-making and enhance business performance.


#  Excel Automation and KPI Dashboard

# This Excel project features a well-organized system of sheets to automate data processing, generate KPIs, and provide insights efficiently. The core functionalities are distributed across various sheets, each serving specific purposes:

#  Sheets - (Calculated Fields, Department_KPIs, Agent_wise, Daily_KPI):
These sheets are fully automated using advanced Excel formulas. They incorporate data validation techniques, conditional formatting, and statistical methods to derive meaningful insights. The KPIs on these sheets are dynamically generated based on the raw data available in other sheets. The automation in these sheets allows tasks that typically take hours to be completed in just minutes.

#  Sheets - (RingCentral - Office, Gate Statistics, Agent Activity):
These contain raw data that is crucial for automation. Advanced Excel formulas are applied directly in these sheets to automate data processing, enhancing efficiency and accuracy.

#  Sheets - (Sales Returns, Y2Y Sales, Inventory Lookup, Emp.Lookup, Audits, WO & Quote):
These sheets solely consist of raw data without any automation applied. The data in these sheets is extracted from various platforms and serves as a source for generating insights and KPIs.

# Sheet - (Links):
This sheet functions as a navigation tool, allowing users to switch quickly to the desired sheet within the same workbook. It also contains URLs that link to the platforms or sources from which the data is derived.

#  Sheet - (Agents):
This sheet provides a categorized list of agents by department type. It helps in segmenting and analyzing agent performance and data based on their respective departments.

#  Overall, this project is designed to streamline data analysis, automate reporting, and provide actionable insights with minimal manual intervention.
