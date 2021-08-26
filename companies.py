import os

import openpyxl as op

SESSION_NAME = 'Technical Day TWO: Engineering and IT - Wednesday, Sep 15, 9:00 am - 2:00 pm EDT' #CHANGE THIS LINE!
#'Technical Day ONE: Engineering and IT - Tuesday, Sep 14, 9:00 am - 2:00 pm EDT'
#'Technical Day TWO: Engineering and IT - Wednesday, Sep 15, 9:00 am - 2:00 pm EDT'
#'Professional Day: Business and Arts and Sciences - Monday, Sep 13, 9:00 am - 2:00 pm EDT'

'''
Contains Company class with relevant (and maybe some irrelevant) company info

getSortedCompanies() takes in a full path to an Excel workbook filename as input, opens said workbook, stores relevant company
information, filters out companies not attending the current session, and then creates a sorted list with the following order:
    Premium Booths, Electricity Booths, Big Companies, then all others.

NOTE: 
    The constant SESSION_NAME (line 5) must be initialized with one of the session options (comments on lines 6-8)
    For future use, options on lines 6-8 must be updated to new session names in Handshake (which writes to the registration Excel file)
    The following column names in the Excel file must not change:
        Employer
        Sessions
        Employer Industry
        Requested Booth Options
        Combined Majors
        Electric
        Big Company
    If the column names do change, change the relevant column name to the updated version on approximate lines 56-62
'''

class Company:
    def __init__(self, employer, sessions, industry, booths, majors, needsElectric, bigComp):
        self.employer = employer
        self.sessions = sessions
        self.industry = industry
        self.booths = booths
        self.majors = majors
        self.needsElectric = needsElectric
        self.bigComp = bigComp
    def printCompanyInfo(self):
        print('-------------------------------------------')
        print('Employer: ', self.employer)
        print('Sessions: ', self.sessions)
        print('Industry: ', self.industry)
        print('Booths: ', self.booths)
        print('Majors: ', self.majors)
        print('Needs Electricity?: ', self.needsElectric)
        print('Big Company?: ', self.bigComp)
    
def getSortedCompanies(workbookName):
    wb = op.load_workbook(workbookName)
    sheet = wb.active
    # determine which column number each relevant data is stored in
    for col in range(1, sheet.max_column + 1): 
        currCell = sheet.cell(row=1, column=col).value
        if currCell == 'Employer': employerCol = col
        if currCell == 'Sessions': sessionsCol = col
        if currCell == 'Employer Industry': industryCol = col
        if currCell == 'Requested Booth Options': boothsCol = col
        if currCell == 'Combined Majors': majorsCol = col
        if currCell == 'Electric': electricCol = col
        if currCell == 'Big Company': bigCompCol = col

    currSessComps = []

    # initialize and store all companies attending the current session
    for row in range(1, sheet.max_row + 1):
        if SESSION_NAME in sheet.cell(row=row, column=sessionsCol).value:
            currSessComps.append(Company(
                employer=sheet.cell(row=row, column=employerCol).value, 
                sessions=sheet.cell(row=row, column=sessionsCol).value,
                industry=sheet.cell(row=row, column=industryCol).value, 
                booths=sheet.cell(row=row, column=boothsCol).value, 
                majors=sheet.cell(row=row, column=majorsCol).value,
                needsElectric=bool(sheet.cell(row=row, column=electricCol).value),
                bigComp=bool(sheet.cell(row=row, column=bigCompCol).value)))

    premBooths = []
    elecBooths = []
    bigComps = []

    # get list of all companies needing premium booths, electric booths, and big companies
    for comp in currSessComps:
        # determine if company needs premium booth for this session
        if len(comp.booths.split(', ')) > 1: # check if they need more than 1 booth (AKA if they're going to more than 1 session)
            boothCount = 0 # index of the list of booths
            booths = comp.booths.split(', ')
            # increment boothCount until the correct index is found (i.e. the one corresponding to the current session)
            for sess in comp.sessions.split('; '):
                if sess == SESSION_NAME:
                    currSessBooth = booths[boothCount] 
                    break
                else:
                    boothCount += 1
            # found the booth for current session- append it to premBooths if it's a Premium Booth
            if currSessBooth == 'Premium Booth':
                premBooths.append(comp)
        else: # company only need 1 booth (i.e. only going to one session)
            if comp.booths == 'Premium Booth':
                premBooths.append(comp)
        # store companies that need electric
        if comp.needsElectric:
            elecBooths.append(comp)
        # store companies that are considered 'Big Companies'
        if comp.bigComp:
            bigComps.append(comp)
    
    # sort current session's companies by premium booths, then electric, then big companies, then just append the rest
    sortedComps = []
    for comp in premBooths:
        sortedComps.append(comp)
    for comp in elecBooths:
        if comp not in sortedComps: sortedComps.append(comp)
    for comp in bigComps:
        if comp not in sortedComps: sortedComps.append(comp)
    for comp in currSessComps:
        if comp not in sortedComps: sortedComps.append(comp)

    return sortedComps


if __name__ == "__main__":
    comps = getSortedCompanies('C:\\Users\\Robbie\\Documents\\Documents\\Tribunal\\Fall 2021\\registered 8-11.xlsx')
    for comp in comps:
        print(comp.employer)
