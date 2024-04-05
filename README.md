Excel Sheet Locker

Excel Sheet Locker is a VBA program that enables password protection and access control for individual sheets within an Excel workbook. It provides a secure way to hide sensitive data and restrict access to specific tabs.
Features:
•	Assign passwords to individual tabs
•	Hide tabs from view and protect them with passwords
•	Unlock specific tabs by entering the corresponding password
•	Admin password feature to unlock all tabs at once
•	Toggle visibility of UI elements (ribbon, formula bar, status bar, workbook tabs)
•	Navigate to the "Home" tab
•	Exit the workbook from the "Home" tab

Tab Structure:
The workbook consists of the following tabs:
•	Home: Contains buttons for accessing the program's macros and functions.
•	Department tabs (e.g., Marketing, Product Development, Customer Service, Finance): Each tab contains department-specific data.
•	Password Key: Contains the passwords for each tab.

Usage:
1.	Open the Excel workbook containing the Excel Sheet Locker program.
2.	Navigate to the "Home" tab to access the program's buttons and functions.
3.	Click on the desired button to perform a specific action:
   Lock Tabs: Hides all tabs and protects them with their assigned passwords.
   Unlock Sheet: Prompts for a password to unlock a specific tab.
   Admin Unlock All Tabs: Prompts for the admin password to unlock all tabs at once.
   Toggle UI Visibility: Hides or shows the ribbon, formula bar, status bar, and workbook tabs.
   Exit: Closes the workbook.

   
Administration Functions:
The program includes two administration functions accessible from the "Home" tab:
1.	Admin Unlock All Tabs: Allows an administrator to unlock all tabs by entering the admin password (default: "Admin01").
2.	Toggle UI Visibility: Enables an administrator to hide or show the ribbon, formula bar, status bar, and workbook tabs by entering the start/stop password (default: "Start01").
                Note: Modify the AdminUnlockAllTabs and ToggleUIVisibility subroutines to change the admin and start/stop passwords, respectively.

Installation:
1.	Open the Excel workbook where you want to use the Excel Sheet Locker program. (Or download the demo excel file)
2.	Press Alt+F11 to open the Visual Basic Editor (VBE).
3.	In the VBE, go to "Insert" > "Module" to create a new module.
4.	Copy and paste the provided VBA code into the new module.
5.	Save the workbook as a macro-enabled workbook (*.xlsm).
   
Disclaimer:
Excel Sheet Locker is intended for educational, demonstration, and personal use only. While it provides a basic level of protection, it should not be relied upon as a sole security measure for sensitive data. 
Always ensure proper security practices and consider using additional security measures when dealing with confidential information.

