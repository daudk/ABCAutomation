# ABC Automation Platform
## 1.Introduction
This is a user-interface based platform built in Python 3.6.6, depending largely on the tkinter (GUI), pandas (data handling), pywinauto (windows automation) and selenium (web automation) libraries. The platform is set up to interact with internal Canon tools like S21, Oracle and IDS for S21 as well as with local operations and processes, especially spreadsheet related operations.
The intended users are ABC team members on Windows based machines (pywinauto will fail on other OSs). 

## 2. Modules
There are 4 main modules that divide the platform into 4 broad use cases that sometimes overlap. This is done mainly to maintain structure and make the menus easie rto navigate.

### 2a. IDS, S21, Oracle Module:
This module houses functions that interact with Canon's proprietary systems and extracts data directly from those systems. All operations require user to enter creditionals which only exist while that specific function is running and never persist. Will consider moving to having user enter credentials directly in browser.

### 2b. SOX Scoping Module:
The SOX Scoping module contains mostly extremely specific SOX related procedures. Can pull and rec out reports, create populations and extract samples etc.

### 2c. Spreadsheets Module:
Module contains some simple spreadsheet operations like merging files, renaming files, extracting samples from population etc.

### 2d. Under Development Module:
This module contains unfinished scripts that are current under production. Users should not use these scripts unless specifically asked to do so, they may be prone to breaking or crashing.

## 3. List of functions

#### Get Accounts Hitting 340
Located under [IDS, S21, Oracle Module](#2a.-IDS,-S21,-Oracle-Module:)
