# Baccalaureate Graduates Survey Response Rate Report

A Python Jupyter Notebook which generates an Excel worksheet of response rates for the BC Student Outcomes Baccalaureate Graduates Survey

# Usage
1. Install Pandas: pip install pandas

2. Install Numpy: pip install numpy

3. Update the value of 'raw_data_path' to the response rate raw data

4. Run each cell in the Notebook one at a time

5. Look at the results in Summary report tables.xlsx

Note: this program will look for hard coded column names to extract the data from. Make sure that your raw data has these exact column names:

* INST (institution acronym eg: UBC, SFU)
* INST_NAME (institution full name)
* PROGRAM (name of the program or the institution respondent total)
* Gross Frame (total respondents)
* Survey Completes (total completed surveys)
* Gross Response Rate (Total response rate)
* Telephone Completes (Total telephone survey completes)
* Telephone Response Rate (Total response rate for telephone surveys)
* Web Completes (total response rate for web surveys)
* To Reach Target (target respondent total)
* Cases Still in Calling Queue (number of cases vendor can still call)

Make sure that the raw data contains different rows for institution totals where the PROGRAM column is "INST TOTAL" and one separate row for the grand total where the INST_NAME column is "Grand Total"

Eg:

| INST        | INST_NAME   | PROGRAM  |
| ----------- | ----------- | -------- |
| Grand Total |             |          |
| BCIT        | British Columbia Institute of Technology|INST TOTAL|


