# People-Operations-JetBrains-Internship-Application
Repository created for internship application purposes.



1\. **Creating an Excel spreadsheet**
I did the task in the online version of Microsoft Excel. I started the task by defining the column names, as requested, I named them:

Employee name, Start date, Position, Type of contract

Then I inserted fictitious first and last names as data for five employees, and that way, we got a realistic table. Then I converted the data into a Table (insert + table). Also, I formatted the table as follows:

Bold header row, set date format (Date format – DD/MM/YYYY). 

At one point, Excel was not properly recognizing dates because they were treated as text. I solved it by checking the column format and converting the values ​​to the correct Date format.

2\. **Column with formula**
I added a new column called: "Employment period (until 31.01.2026)"

For the calculation, I used the formula with the function DATE(2026;1;31) to define the end date. Then I calculated the difference between the Start date and 31.01.2026 using the appropriate date calculation function. I tried YEARFRAC first, but I didn't like how the data looked because it was in decimal, so I used DATEDIF. The final formula looked like this:

=DATEDIF(B2; DATE(2026;1;31);"d")

The table was formatted as an Excel Table, so the formula was automatically applied to all rows, and with that, I enabled the employment period to be dynamically calculated for each employee.
<img width="926" height="149" alt="image" src="https://github.com/user-attachments/assets/ae06b553-bbd4-4a16-8f11-abd7d4f78c1f" />

3\. **Creating an Employment Confirmation template in Word**

I used the Microsoft Word 2010 application, and I formatted the document in the simplest way. That part could be improved, but I didn't focus on it that much. I would insert all the contact details of the person who needs to sign the document and the company logo. I set the date at the beginning of the document to update automatically. At the end of the document, I inserted a line for signature and a stamp.

For automation, I used Mail Merge functionality (Mailings + Start Mail Merge + Letter)

I linked the document to the Excel spreadsheet via Select Recipients + Use Existing List

I added following Merge Fields:

«Employee name», «Start date», «Position», «Type of contract», «Employment\_period\_until\_31012026»

After that I used the Preview Results option to check the correctness of the data. Finally, I clicked Finish \& Merge, which automatically generated individual receipts for all employees from the spreadsheet.

The final document now looks like this:
Date: 2/14/2026

To whom it may concern,

This is to confirm that «Employee\_name» has been employed with our company since «Start\_date». The employee currently holds the position of «Position» under a «Type\_of\_contract» contract. As of January 31, 2026, the total period of employment amounts to «Employment\_period\_until\_31012026» days. This confirmation is issued upon the employee’s request for official purposes.

Should you require any further information, please do not hesitate to contact us.

Sincerely,

&nbsp;

HR Department

JetBrains


<img width="918" height="751" alt="image" src="https://github.com/user-attachments/assets/11613257-e66e-496b-848c-56e78b3481f1" />



4\. **Proposal for additional automation**



As an additional automation of the process, I would suggest the introduction of a digital request through Google Forms, where employees who want Employment Confirmation would fill out a standardized form in which they would enter their data and necessarily an email address so that the document could be delivered directly to them. The link to the form would be sent through a circular email to all employees with a brief explanation of the procedure, which would eliminate individual manual requests via email. Each completed request would be automatically recorded in a Google Sheets table, thus creating a centralized database that serves as a data source for further processing. That file could then be downloaded as .xlsx and used to calculate the employment period through a formula, after which the Mail Merge process in Word would be applied to automatically generate a confirmation. If the organization uses digital signatures, the confirmations could be automatically converted to PDF and sent to employees via email, which would automate the entire process from the request to the delivery of the document as much as possible. I also assume there are apps that do that.


