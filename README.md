# ABC Automation Platform
## 1.Introduction
This is a user-interface based platform built in Python 3.6.6, depending largely on the tkinter (GUI), pandas (data handling), pywinauto (windows automation) and selenium (web automation) libraries. The platform is set up to interact with internal Canon tools like S21, Oracle and IDS for S21 as well as with local operations and processes, especially spreadsheet related operations.
The intended users are ABC team members on Windows based machines (pywinauto will fail on other OSs). 

## 2. Modules
There are 4 main modules that divide the platform into 4 broad use cases that sometimes overlap. This is done mainly to maintain structure and make the menus easie rto navigate.

### 2a. IDS, S21 and Oracle Module:
This module houses functions that interact with Canon's proprietary systems and extracts data directly from those systems. All operations require user to enter creditionals which only exist while that specific function is running and never persist. Will consider moving to having user enter credentials directly in browser.

### 2b. SOX Scoping Module:
The SOX Scoping module contains mostly extremely specific SOX related procedures. Can pull and rec out reports, create populations and extract samples etc.

### 2c. Spreadsheets Module:
Module contains some simple spreadsheet operations like merging files, renaming files, extracting samples from population etc.

### 2d. Under Development Module:
This module contains unfinished scripts that are current under production. Users should not use these scripts unless specifically asked to do so, they may be prone to breaking or crashing.

## 3. List of functions

#### Get Accounts Hitting 340
Located under IDS, S21, Oracle Module.
Given a subsidiary and year, will download the 340 report for the sub and year to a directory and will then extract a new file called output.xlsx that keeps only the Account Number and Total Actuals column and adds a new column called size ranging from 0 to 3 depending on how much money flowed through a given account in a year.
Useage:
- Clicking the button will open a file browser dialog, navigate to and select any directory. 
- Will then ask for subsidiary, please enter full name of report e.g "CUSR0340".
- Enter year as a 4 digit number e.g 2017
- Enter Single Sign On Username and Password
- Click Ok
- New Chrome Browser should open and the program will freeze until it is done.
##### Note: This is a very specific function that not many people will need. If unsure about capability or its use, contact developer.

#### Download All
Located under IDS, S21, Oracle Module.
Will save all reports that are ready for download in the "View Reports" section of your IDS.
Useage:
- Clicking the button will open a file browser dialog, navigate to and select any directory. 
- Enter Single Sign On Username and Password
- Click Ok
- New Chrome Browser should open and start downloading your reports in .xlsx formar the program will freeze until it is done.

##### Note: Please wait until all reports have finished running in IDS otherwise behavior may be unpredictable.
##### Note: If file is not available in .xlsx format, it will default to ztab format. This behavior can be reconfigured, contact developer.


#### Run 315 Reports
Located under IDS, S21, Oracle Module.
Will run the 315 report in bulk for a given subsidiary for a given year.
Useage:
- Clicking on button will reveal a number of options.
- Enter account numbers in the whit space in the left, each account must be separated by a new line. e.g:
    11711005
    21211001
    65301205
- In the Subsidiary drop down, choose a subsidiary, default option will alsways be CUSA
- In the next drop down choose the year, default will always be current year - 1.
- The frequency drop down provides an option not available within IDS. For larger accounts, all the transactions from 1 year may not fit in 1 report so you can specify how many reports to break the request into. Yearly will pull 1 report for full year, Semi-annually will pull 2, quarterly will pll 4 per account and monthly will pull 12 per account. Default is set to yearly.
  ##### Note: If unsure, leave this option set to yearly
- Choose source. Default is set to "All Values". 
  ##### Note: The Source drop down displays all options available in IDS, some options may not exist for certain subs. Verify the source actually does exist for the sub you are requesting.
- Program will become unresponsive while the reports are being requested but you should be able to view progress in the Chrome window that pops up. Avoid interacting with the Chrome window in any way while program is running, this may lead to unpredictable behavior.
##### Coming soon: Multithreading to allow view of progress bars and so that program does not become unresponsive while one thread is running.

#### Expense Reports
Located under IDS, S21, Oracle Module.
Will take screenshots of expense reports overview page and download all attachements and neatly organize them into 1 folder per expense report.
Useage:
- Enter the expense report numbers in the white space to the left. Each expense report number should be separated by new line e.g:
    SSE1257674
    SSE1254363
    SSE1259299
    SSE1234148
    SSE1257589
    SSE1259492
- Specify download directory by clicking Browse button. Default will be C:\Users\<user-name>\Downloads\
- Click Go button, program will become unresponsive but progress can be followed in the Chrome window that pops up. Avoid interacting with the Chrome window in any way while program is running, this may lead to unpredictable behavior.
##### Coming soon: Multithreading to allow view of progress bars and so that program does not become unresponsive while one thread is running.
