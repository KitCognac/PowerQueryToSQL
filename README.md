# PowerQueryToSQL
Using Power Query in Excel as a SQL Server Import Function

## Main Concept
Using Power Query Function to Create bunch of SQL Queries then send them to SQL Server to import data.

## Power Query Setting
Config 2 settings in Query Options ( in Excel, Ribbon Data => Get Data => Query Options => GLOBAL or Current Workbook depends on your needs )
- In Data Load, tick "Never allow data previews to download in background". If this setting is not off, Import Queries will be sent to SQL Server automatically control by some trigger in Excel.
- In Security, uncheck "Require user approval for new native database queries". By default, Excel will ask for permission everytime Native Query changed. Sending 100 SQL Insert Command Excel will ask 100 times, we need to turn that setting off

## For ADVANCED USER only
This function will not replace SSIS or any kind of insert bulk, but it will: 
 - Ultilize settings per table in excel workbook through Queries&Connection Task Pane. 
 - Easy maintain DML&DDL before upload. Having a coffee or delete a table in database, choose wisely.
 - Upload directly while Excel Workbook is opened.

## Limitation
 - Unexpected import Data when user hover queries in Queries&Connection Task Pane
 - Slow compared to insert Bulk because rely on SQL Server implicit conversion from string and multiple SQL Files Reading.

## Usage
- This function works best and fast for importing data under 100k rows.
- Create and bind a User-Defined Parameter to IS_START Parameter of this function. After import job are done, set that parameter to NO so that all import job can not import any data when hover over queries in Queries&Connection Task Pane or when PQ trying to download preview.
- Please cross-check after import and report any strange phenomenon like duplicated imported data or DDL&DML not exec.
