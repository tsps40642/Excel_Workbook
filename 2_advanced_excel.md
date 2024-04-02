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

### Format the PivotChart
When we sort the data in the table, its PivotChart has inverse order since the table and the y-axis are inverse.   
We can solve this problem by formatting axis:   
Right click the y-axis of the PivotChart -> Format Axis -> select "Categories in reverse order"  

### Other edit tips 
At the PivotChart:  
1. Hide field bottons: Right click the field bottom -> Hide All Field Bottons on Chart
2. Click "+" on the top right to select Data Labels and other things you want  

In the table:   
1. Add dollar sign: Right click the data -> Number Format -> Currency
2. In the Design tab -> Report Layout -> Layout -> Show in Outline Form, to have column name on the table, and other layout you want to add

## Change data source
Remember to create the table for the data that you also want to have a pivot table:  
1. Click a cell + control T to create a table
2. In the Table Design tab -> rename the table
3. In PivotTable Analyze -> Change Data Source -> enter the table name we just built, so that you don't have to build anothr pivot table

## Going into details
Double click the value that you want to look at its details -> it will generate a table -> remane the sheet 

## Data Validation
Data validation make sure the data being input is correct.  
1. Select the whole column of Sales Rep -> in the Data tab -> Data Tools -> Data Validation -> determine the Validation criteria (Allow, Source)
2. Can enter the title and message for the users in the Iuput Message
3. Can enter the title and message for the users in the Error Alert
4. If want to change the value, one should change in the source we just select 

## Conditional formats
1. In the Home tab -> Conditional Formatting -> select the format you want
2. Can also manage the rule by selecting Manage Rules
3. Can also use Icon Sets -> then Manage Rules to edit it if you need
4. Select the column name row -> in Data tab -> Filter: this can let you sort by the conditional format we just set

## Comments and notes
One can add comments or notes in a cell.  

## Link the cell 
1. In the same sheet in cell_2, type "= (select cell_1)" to link cell_1 and cell_2
2. In the different sheet in sheet_2 cell_2, type "=Sheet_1!(select cell_1)" to link sheet_1 cell_1 and sheet_2 cell_2

## Privacy and Protection 
### Unlock cells you want uers to have access to
Select the cells want to unlock -> right click -> Format Cells -> Protection -> unselect "Locked"

### Protect the worksheet i.e. enables the protection
In the Review tab -> Protect -> select the fields that you allow users to access
