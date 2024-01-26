# Excel Projects
## About
This contains completed Excel projects that passed through cleaning and manipulation following several steps including Data Quality Assessment, Reshaping Data, Handling Missing Values, Descriptive And Relational Analysis
# Data Cleaning
## Importing Data 
- Import source data tables into separate sheets within an Excel workbook from CSV files or external database sources using the Get & Transform Data tool 
## Data Quality Assessment
- Scan for blank cells, invalid entries, and duplicate records using functions like COUNTIF, ISBLANK, Conditional Formatting
- Analyze distributions of numeric columns with Descriptive Statistics under the Data Analysis toolkit to identify outliers
## Reshaping Data
- Split columns with multiple attributes into separate columns using the Text to Columns tool
- Group worksheets with related data tables using Sheet References to link cell ranges
## Handling Missing Values
- Remove rows with blank key column values if negligible missing data  
- Set default placeholders for date/numeric cells using IFERROR
- Fill blank cells by interpolating from neighboring rows with Flash Fill 
## Standardization 
- Trim and clean trailing spaces using TRIM
- Conform date layouts to consistent YYYY-MM-DD format with DATEVALUE 
- Validate and codify text values with Data Validation dropdown lists
## Deduplication
- Identify duplicate entries based on composite key columns with COUNTIFS
- Flag or filter out dupes; keep the most recent record with the Remove Duplicates tool
- 
# Data Exploration
## Descriptive Analysis
- Aggregate and transform data with Pivot Tables and Charts to summarize distributions 
- Enable subtotals for grouping numbers by categories in Pivot Tables  
## Relationship Analysis 
- Compute correlation coefficients between numeric variables using the Data Analysis Correlation tool
- Quantify effect direction and strength to discover linear models 
