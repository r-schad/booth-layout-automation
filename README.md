# booth-layout-automation

The constant sessionName (line 5) must be initialized with one of the session options (comments on lines 6-8)
For future use, options on lines 6-8 in companies.py must be updated to new session names in Handshake (which writes to the registration Excel file)
The following column names in the Excel file must not change:
    Employer
    Sessions
    Employer Industry
    Requested Booth Options
    Combined Majors
    Electric
    Big Company
If the column names do change, change the relevant column name to the updated version on approximate lines 70-76 in companies.py
Additionally, a column must be added directly after the last column in the excel file. This column must be titled "Big Company".
All "Big Companies" (as decided by Career Dev) should be marked with a "1" in the Big Company column
