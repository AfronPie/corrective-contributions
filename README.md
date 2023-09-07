# Introduction
This code will do the math for corrective contributions. A typical situation may sound like the following:
A 401(k) participant named George elects to have $300 withheld from his paycheck and placed in his 401(k) every two weeks. However, his employer's payroll system had a
error - which caused the $300 to be withdrawn from his 07/14/23 pay but not deposited into his 401(k) account. It is now 08/20/23, and George's employer has discovered this error. George's employer immediately deposits $300 into George's 401(k) account to account for the mistake. However, the George's overall portfolio has gone up 5.8% in the time period between 07/14/23 and 08/20/23. How can we accurately account for this increase in value? 

# Installation
Aside from the native python libraries - you will need [pandas](https://pypi.org/project/pandas/) and [docx](https://pypi.org/project/python-docx/).

# Process
The process to use will involve the following:
1. Make a copy of the Corrective Contributions .ipynb template, along with the Funcs.py python file. 
2. Create an excel sheet which contains your information, using the Data Import Template.xlsx file as a guide on how to format your information. Once you have created the sheet (note that the column names and information should exactly match the Data Import Template.xlsx), convert the excel file to .csv - which will allow you to easily read it into the .ipynb file.
3. Run the cells in the .ipynb file in a jupyter notebook (google collab works too). Overall, the input is 1 csv file - which outputs 1-2 excel files and 1-2 word document files depending on the circumstances of the case.

# Specific Abbreviations Used
* EE = Employee
* ER = Employer
* Pay Date = The date funds were supposed to have been deposited into EE’s account
* VTD Date = The actual deposit date into EE’s account
* VTD will always be after the Pay Date (or we wouldn’t have a corrective contribution case)
