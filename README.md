# Excel-VBA-Sales-Automation:

Transferring data in Excel between cells to cells, sheet to sheet, workbook to workbook can make the process of completing tasks both more time consuming and more error prone. This is a common source of frustration that plagues the likes of data entryists, data anlysts, supply chain operators, economics, accountants, you name it, across the realm of corporate. If your day to day Excel related tasks or at least a portion of them are perfectly replicable, it is time to greet the 21st century with open arms and automate said tasks. Automating tasks is a tall order for an employee if programming is alien to them. That is where this repository comes in (fingers crossed). This respoistory contains four Excel VBA programs (formally referred to as "Macros") that provides a foundation for how onne can automate the process of entering, cleaning, and transferring data in Excel for corporate operations. 

# Brief overview of each of the four macros:

Sub Block1(): 
Task 1: Deletes unwanted entities in the raw data that is exported from the company's Shopify sales database into Excel and transferred over to the user's Macro Enabled Workbook.
Task 2: Scans the data for potential duplicated orders.
Task 3: Scans the data for missing orders that slipped under the user's radar while they were exporting orders from Shopify's database, so the user can rectify this man made error. 
Task 4: Deletes columns storing unnecessary data, formats the cell contents that caters to the user's needs. Now the user can finish some very small finishing touches that have to be completed manually before the user utilizes the next macro to transfer the refined data into the destination worksheet of the workbook.

Sub Block2(): 
Task 1: Transfers the raw data into the destination worksheet of our macro enabled workbook. A portion of the data will be first transferred to intermediate worksheets to be further refined.
Task 2: Clears the workbook of no longer necessary so there would be space in the workbook for the next batch of sales order data to be processed for the day. 

Sub Block3(): 
Task 1: Transfers the raw data in our macro enabled workbook to a destination workbook that has the hyperlinks needed to import the sales order data into the company's Enterprise Resource Planning (ERP) application.
Task 2: Clears the contents of the marco enableld workbook then saves and closes BOTH the macro enabled workbook and the destination workbook while the user can spend the next several minutes sending emails to their colleagues and working with the aforementioned ERP application. 

Sub Block4(): 
Task 1: Keeps record of the Shopify sales order data the user processed for the day in a record keeping worksheet of the macro enabled workbook. 
