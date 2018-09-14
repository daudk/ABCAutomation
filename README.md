# ABC Automation Platform
## 1.Introduction
This is a user-interface based platform built in Python 3.6.6, depending largely on the tkinter (GUI), pandas (data handling), pywinauto (windows automation) and selenium (web automation) libraries. The platform is set up to interact with internal Canon tools like S21, Oracle and IDS for S21 as well as with local operations and processes, especially spreadsheet related operations.
The intended users are ABC team members on Windows based machines (pywinauto will fail on other OSs). 

## 2. Modules
There are 4 main modules that divide the platform into 4 broad use cases that sometimes overlap. This is done mainly to maintain structure and make the menus easie rto navigate.

### 2a. IDS, S21, Oracle Module:
This module houses functions that interact with Canon's proprietary systems and extracts data directly from those systems.
