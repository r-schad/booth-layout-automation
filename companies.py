import os

import openpyxl as op

class Company:
    def __init__(self, employer, industry, premiumBooth, majors, sessions, needsElectric):
        self.employer = employer
        self.industry = industry
        self.premiumBooth = premiumBooth
        self.majors = majors
        self.sessions = sessions
        self.needsElectric = needsElectric
    def printCompanyInfo(self):
        print('-------------------------------------------')
        print('Employer: ', self.employer)
        print('Industry: ', self.industry)
        print('Premium Booth?: ', self.premiumBooth)
        print('Majors: ', self.majors)
        print('Sessions: ', self.sessions)
        print('Needs Electricity?: ', self.needsElectric)
    
if __name__ == "__main__":
    wb = op.load_workbook('C:\\Users\\Robbie\\Documents\\Documents\\Tribunal\\Summer 2021\\All_Registrants.xlsx')
    sheet = wb.active
    # Get which column number each relevant data is stored in
    for col in range(1, sheet.max_column + 1):
        currCell = sheet.cell(row=1, column=col).value
        if currCell == 'Employer': employerCol = col
        if currCell == 'Employer Industry': industryCol = col
        if currCell == 'Requested Booth Options': premiumCol = col
        if currCell == 'Combined Majors': majorsCol = col
        if currCell == 'Sessions': sessionsCol = col
        if currCell == 'General Items - Access to Electric': electricCol = col
    companies = []
    for row in range(1, sheet.max_row + 1):
        companies.append(Company(
            employer=sheet.cell(row=row, column=employerCol).value, 
            industry=sheet.cell(row=row, column=industryCol).value, 
            premiumBooth=sheet.cell(row=row, column=premiumCol).value, 
            majors=sheet.cell(row=row, column=majorsCol).value,
            sessions=sheet.cell(row=row, column=sessionsCol).value, 
            needsElectric=sheet.cell(row=row, column=electricCol).value))
    
    for comp in companies:
        comp.printCompanyInfo()
