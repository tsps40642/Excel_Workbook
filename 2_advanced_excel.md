# Advanced_Excel_Workbook
This is a repository for some useful Excel notes.  

## Well-defined list
1. Headers are unique i.e. no duplicate columns
2. Headers and data are indep. and unique i.e. no data in the header 
3. Complete dataset i.e. no blank even if they are hidden

## Create a table
1. Insert -> click Table
2. Or use control + T to create a table on the existing list
3. Rename the table -> then can go to namebox to select the table you want

## Slicer
1. In the Table Design tab -> Insert Slicer -> select the column you want to use as a slicer
2. Can multi-select or unselect

## Total row
In Table Design tab -> Total Row -> will add a row at the bottom of the table that you can customize 

## Remove duplicates 
1. In Table Design tab -> Remove Duplicates -> select the columns that you want to investigate duplicates
2. Then it will automatically show the number of duplicated rows removed and unique values remained

## Flash fill 
1. Fill the first blank to show the rule for flash fill 
2. In Data tab -> Data Tools -> Flash fill, it will automatically find the pattern and fill the blank 
3. You can use control + E on the blank to activate the flash fill 

## Multiple level sort 
1. In Data tab -> Sort & Filter -> Sort -> select the column and sort on as the reference for sorting
2. In Data tab -> Sort & Filter -> Sort -> Add Level for multiple columns sorting

## Subtotal 
1. Important to first sort your columns or you won't get the correct subtotal! 
2. In Data tab -> Outline -> Subtotal: when hitting a difference in a column, it will calculate a function
3. Select the column you want subtotla in "each change in", and select the columns you want to calculate subtotal in "Add subtotal to"
4. There will be different level for the sheet now 

## Quick analysis 
Control + A to select the whole list/table -> click Quick Analysis on the bottom right 

## Charts
### Create a chart and update the data 
1. In Insert tab -> Recommended Chart to pick the one you want to use
2. If later you add new columns or rows into the original list/table, the chart already created won't include newly added
3. To include the new added data: in Chart Design -> Select Data to reselect

### Change chart type 
In Chart Design -> Change Chart Type -> Combo, it allows you to select type for each column (I select barchart for Week 1-4 and line for Sum)   

### Edit the chart
Further edit legend, title, and columns by clicking "+" on the top right

### Save the chart as a template 
So that we can use our own template and keep charts consistent, don't need to edit its look again.  
1. Right click the chart -> Save as Template...
2. The next time when you want to use your template: in Insert tab -> Recommended Chart -> All charts -> Template

## Sparkline 
It shows the line mimicing the data s.t. you can see the trend of it. There are 2 ways to do this:  
### In the Insert tab 
1. Select the row you want to investigate -> in Insert tab -> Sparklines
2. Data Range (already select), Location Range (the place you want to put that sparkline)
3. Can copy to other rows by dragging down on the bottom right (auto fill)

### In the Quick Analysis 
1. Select the row you want to investigate -> in Insert tab -> Quick Analysis -> Sparkline, then select the form you want
2. If having negative value -> can use Win/Loss s.t. it will mark the negative automatically 

### Edit the sparkline 
Select the sparkline you want to edit -> Sparkline tab -> Marker Color -> can mark Low Point, High Point, etc. 

## Pivot tables
### Create a PivotTable
1. Click a cell + control T to create a table
2. In the Table Design tab -> rename the table
3. Tools -> Summarize with PivotTable (and since we already named the table so Table/Range will select this table)
4. Choose New Worksheet to place the PivotTable
5. Rename the created PivotTable
6. Can customize the setting of PivotTable Fields: change columns order, chance display, etc.

### Add a slicer
1. Drag Salesperson into "Rows" or "Columns"
2. Drag Revenue into "Values"
3. In PivotTable Analyze -> Filter -> Insert Slicer -> select Salesperson to create its visual slicer 

### Add a timeline 
1. Make sure you have columns representing time in your data
2. In PivotTable Analyze -> Filter -> Insert Timeline -> select Order Date

### PivotChart 
It's the visualization of the report.  
In PivotTable Analyze -> Tools -> PivotChart -> select the chart form you want 

## 
