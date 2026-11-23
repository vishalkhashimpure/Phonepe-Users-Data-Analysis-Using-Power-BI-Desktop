# Phonepe-Users-Data-Analysis-Using-Power-BI-Desktop

Project Overview
The PhonePe Dashboard project focuses on analysing transaction data using Power BI. The main objective is to understand transaction patterns, payment statuses, service-wise revenue, user activity, age groups and transaction trends.

In this project, I used Power Query for data cleaning and transformation, DAX for creating calculated measures, and Power BI visuals to build an interactive dashboard.

Before starting the analysis, I first explored the dataset to understand the available tables, columns and information.

Step 1: Data Extraction, Transformation and Cleaning
1. Open Power BI Desktop and select Get Data → Excel Workbook. Choose the dataset and click on Transform Data to open Power Query Editor.
2. After loading the data into Power Query Editor, go to the View tab and enable Column Quality, Column Distribution and Column Profile to understand the data.
3. Check all the column headers and make sure there are no spelling mistakes or unwanted spaces in the column names.
4. Check the data types of all columns and make sure each column has the correct data type.
5. Check for duplicate records. If there are unwanted duplicates in a primary key column, select the column, right-click and choose Remove Duplicates. Make sure you don't remove valid transaction records.
6. In the Services column, replace recharge_bills with recharge bills. To do this, select the column, right-click and choose Replace Values.
7. Similarly, replace money_transfer with money transfer.
8. Check for null values, blank rows and incorrect entries. Remove or correct them wherever necessary.
9. Once the cleaning process is completed, click on Close & Apply to load the transformed data into Power BI.
    
Step 2: Creating a Date Table
1. Save the Power BI file with a proper project name.
2. Since the dataset contains a Date column, create a separate Date Table using DAX. This helps in performing time-based analysis such as monthly trends, quarterly analysis and month-over-month growth.
3. Go to the Table View on the left side of Power BI Desktop. Select Table Tools → New Table.
4. Create the Date Table using the following DAX formula:
Date_Table = ADDCOLUMNS(CALENDAR(MIN(All_Transactions[Date]), MAX(All_Transactions[Date])), "Year", YEAR([Date]), "Month No.", MONTH([Date]), "Month Name", FORMAT([Date], "mmm"), "Quarter", "Q" & FORMAT([Date], "Q"), "Weekday", FORMAT([Date], "dddd"), "Day No.", WEEKDAY([Date], 2), "Weekend", IF(WEEKDAY([Date], 2) >= 6, "Weekend", "Weekday"))
5. While creating the Date Table, I included Month No. and Day No. columns. These columns help sort month names and weekdays in the correct order.
6. After creating the Date Table, select the Month Name column and use Sort by Column → Month No. Similarly, select Weekday and sort it by Day No.
7. Now go to Table Tools and select Mark as Date Table. Enable the option and select the Date column as the date column.
8. Save the changes.
   
Step 3: Data Modeling
1. In this step, I created relationships between the different tables to connect the data and make analysis easier.
2. First, create a relationship between the All_Users and All_Transactions tables using the common User_ID column. Drag User_ID from the All_Users table to the User_ID column in the All_Transactions table.
3. Create a one-to-many relationship (1:*) where All_Users is on the one side and All_Transactions is on the many side.
4. If you get an error such as Column contains a duplicate value, and this is not allowed for columns on the one side of a many-to-one relationship, check the User_ID column in the All_Users table.
5. Make sure the User_ID column contains unique values and does not contain blank IDs. Remove only unwanted duplicate or blank records.
6. To remove blank rows in Power Query, select the relevant table and use Remove Rows → Remove Blank Rows, if appropriate.
7. After fixing the issue, create the relationship again.
8. Similarly, create a relationship between Date_Table and All_Transactions using their Date columns. The Date_Table should be on the one side and All_Transactions on the many side.
9. After creating the relationships, create a separate table to store all the DAX measures. Go to Home → Enter Data, provide a table name as Measures and a column name, and click Load.
10. You can hide the temporary column from the Measures table to keep the Fields pane organised.
11. Finally, save the Power BI file using Ctrl + S.
Step 4: Creating KPIs and DAX Measures
In this step, I created DAX measures to calculate important business metrics and track transaction performance.

Select the Measures table and create the following measures using Table Tools → New Measure.

1. Total Transaction Amount

This measure calculates the total amount of all transactions.

Total Transaction Amount = SUM(All_Transactions[Amount])

2. Total Transactions Count

This measure calculates the total number of transactions.

Total Transactions Count = COUNT(All_Transactions[Transaction_ID])

3. Successful Transactions

This measure counts only the transactions with a successful payment status.

Successful Transactions = CALCULATE([Total Transactions Count], All_Transactions[Payment_Status] = "successful")

4. Success Rate

This measure calculates the percentage of successful transactions out of the total transactions.

Success Rate = DIVIDE([Successful Transactions], [Total Transactions Count], 0)

Format this measure as a percentage.

5. Total Users

This measure calculates the total number of unique users.

Total Users = DISTINCTCOUNT(All_Users[User_ID])

6. Transaction Value Previous Month

This measure calculates the total transaction amount for the previous month.

Transaction Value Previous Month = CALCULATE([Total Transaction Amount], DATEADD(Date_Table[Date], -1, MONTH))

7. Transaction Value Month-over-Month Growth %

This measure calculates the percentage growth or decline in transaction value compared with the previous month.

Transaction Value MoM Growth % = DIVIDE([Total Transaction Amount] - [Transaction Value Previous Month], [Transaction Value Previous Month], 0)

8. Total Transactions Previous Month

This measure calculates the total number of transactions from the previous month.

Total Transactions Previous Month = CALCULATE([Total Transactions Count], DATEADD(Date_Table[Date], -1, MONTH))

9. Total Transactions Month-over-Month Growth %

This measure calculates the percentage growth or decline in the number of transactions compared with the previous month.

Total Transactions MoM Growth % = DIVIDE([Total Transactions Count] - [Total Transactions Previous Month], [Total Transactions Previous Month], 0)

10. Previous Month Success Rate

This measure calculates the success rate of the previous month.

Previous Month Success Rate = CALCULATE([Success Rate], DATEADD(Date_Table[Date], -1, MONTH))

11. MoM Success Rate Change %

This measure calculates the relative percentage change in the success rate compared with the previous month.

MoM Success Rate Change % = DIVIDE([Success Rate] - [Previous Month Success Rate], [Previous Month Success Rate], 0)

Format the growth measures as percentages.
Creating Age Segments
1. In the All_Users table, there is an Age column containing different age values.
2. To understand which age groups are using PhonePe more, create a new column called Age Segment.
3. Go to Transform Data → Power Query Editor.
4. Select the Age column and choose Add Column → Conditional Column.
5. Create age groups such as 18–25, 26–35, 36–45, 46–55 and Above 55, depending on the dataset.
6. Click OK and then select Close & Apply to load the changes.
   
After completing these steps, the data cleaning, transformation, data modeling and DAX measure creation are completed. Now the measures can be used to create KPI cards and visualizations.
Creating KPI Cards
Add Card visuals to the dashboard and use the following measures:
- Total Transaction Amount
- Total Transactions Count
- Successful Transactions
- Success Rate
- Total Users
- Transaction Value MoM Growth %
- Total Transactions MoM Growth %
The month-over-month growth KPIs help track how transaction value and transaction count change compared with the previous month. Their values respond to the selected date filters.
Step 5: Creating Visualizations
After creating the required measures, I started building different visualizations to understand transaction patterns and user behaviour.

1. Monthly Transaction Trend – Line Chart

Create a line chart to analyse monthly transaction trends. Put the Month-Year field from the Date Table on the X-axis and Total Transaction Amount on the Y-axis.

2. Revenue by Services – Column Chart

Create a column chart to compare the total transaction amount generated by different PhonePe services.

3. Payment Status Analysis – Stacked Column Chart

Create a stacked column chart to analyse transaction counts based on Payment_Status.

The dataset contains different payment statuses. For analysis, unsuccessful transaction reasons such as Insufficient Amount, Wrong PIN and Server Error can be grouped under a common category called Failed.

To do this, go to Power Query Editor, select the Payment_Status column and use Replace Values or a conditional column to group the relevant statuses.

4. Age Segment Contribution – Donut Chart

Create a donut chart to understand the contribution of different age groups. Use Age Segment as the category and Total Users or Total Transactions Count as the values.

5. Weekend and Weekday Contribution – Donut Chart

Create a donut chart to compare transaction activity on weekdays and weekends. Use Weekend from the Date Table as the category and Total Transactions Count as the values.

6. Daily Transaction Analysis – Column Chart

Create a column chart to analyse the number of transactions on different dates. Use Date from the Date Table on the X-axis and Total Transactions Count on the Y-axis.

7. Top 10 Users – Column Chart

Create a column chart to identify the top 10 users who made the highest number of transactions.

To display only the top 10 users:
1. Select the chart.
2. Open the Filters pane.
3. Select the User Name field.
4. Change the filter type to Top N.
5. Enter 10 in the Show Items section.
6. Select Total Transactions Count under By Value.
7. Click Apply Filter.
The chart will display the top 10 users based on their transaction count.

Step 6: Creating Tooltips
Tooltips help display additional information when hovering over a visual. They make the dashboard more interactive and provide extra details without adding more charts to the main page.
1. To create a tooltip, click the plus (+) icon near the report page tabs to create a new page.
2. Select the blank page and open the Format Page settings.
3. Go to Canvas Settings and change the page type to Tooltip.
4. Under Page Information, enable Allow use as tooltip.
5. Now create a column chart on the tooltip page using Services as the category and Total Transaction Amount as the value.
6. Go back to the main dashboard page and select the visual where you want to connect the tooltip.
7. Open the Format Visual settings and find the Tooltip option. Enable it if it is turned off.
8. Change the Tooltip Type to Report Page and select the tooltip page you created.
9. Now hover over the selected visual to check whether the tooltip is working correctly.
10. Similarly, create another tooltip page and connect it to the Weekday and Weekend donut chart to display additional information.
    
Step 7: Dashboard Design

After completing the data analysis and visualizations, I worked on designing the dashboard to make it more attractive, organised and interactive.
1. First, download the dashboard background image from GitHub and save it on your device.
2. Click anywhere on the blank report page and open Format Page settings.
3. Go to Canvas Background or Wallpaper, click Browse and select the background image from your device.
4. After setting the background image, remove the background and borders of the visuals to make them blend properly with the dashboard design.
5. Select each visual, open Format Visual and search for Background. Turn it off or set the transparency to 100%. Similarly, turn off the Border option.
6. If some formatting options are not visible, expand all categories in the Format Visual pane to access the additional settings.
7. Arrange all the visuals properly on the dashboard background. Keep the KPI cards at the top and arrange the charts below them.
Conditional Formatting for Growth KPIs
8. Select the two month-over-month growth KPI cards and apply conditional formatting to their values.
9. Open Format Visual and go to Callout Values. Select the fx option next to the colour setting.
10. Choose Rules as the formatting style and set the conditions so that positive growth values appear in green and negative growth values appear in red.
11. Apply the same formatting to the other growth KPI.
Adding Slicers
12. Add two slicers to the dashboard:
    - Month Name from Date_Table
    - Payment_Status from All_Transactions
13. These slicers help users filter the dashboard based on the selected month and payment status.
Formatting Visuals
14. To change the colours of the bar and column charts, select the visual and open Format Visual.
15. Go to Columns or Data Colors, depending on the visual and Power BI version, and choose the colours you want.
16. Apply suitable colours to the remaining visuals to maintain a consistent dashboard design.
Changing Slicer to Dropdown
17. Select the Month Name slicer and open Format Visual.
18. Go to Slicer Settings and change the slicer style to Dropdown.
19. Repeat the same process for the Payment_Status slicer if required.
20. Finally, check all the visuals, KPI cards, slicers and tooltips to make sure everything is working properly.
21. Save the completed dashboard using Ctrl + S.
