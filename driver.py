# driver.py
#
# A part of the IOC Assistant software package, developed by Brandon Miller (DHS-CISA Intern 2019-2020)
#
import xlsxwriter
import methods
import bs4, sys, urllib
from datetime import datetime

# Useful variables.
inquiring = True
currentIOC = 2

# Gather name for excel file.
print("Welcome to the IOC Assistant python program, written by Brandon Miller.")
filename = input('Enter your desired file name, leaving off the file suffix (.xlsx): ') + '.xlsx'
print(filename)

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(filename)
ws = workbook.add_worksheet()

# Formatting
bold = workbook.add_format({'bold': 1})
ws.set_column(1, 1, 15)

# Add headers
ws.write('A1', 'Indicator Name', bold)
ws.write('B1', 'IOC Source Report Name and Report Location (Most Important/Also pls defang URIs)', bold)
ws.write('C1', 'Activity Type (i.e. Phishing, etc.)', bold)
ws.write('D1', 'Malware Name', bold)
ws.write('E1', 'Indicator Type (C2, dropper, etc.)', bold)
ws.write('F1', 'Actor Name', bold)
ws.write('G1', 'Associated IP Address (does it currently resolve? Multiple Ips?)', bold)
ws.write('H1', 'Virus Total Ratio/Other Relevant Info', bold)
ws.write('I1', 'Alexa Page Rank', bold)
ws.write('J1', 'Total Number of Subdomains', bold)
ws.write('K1', 'Domain Registrant (Owner) ', bold)
ws.write('L1', 'Domain Registration and Expiration  Date', bold)
ws.write('M1', 'Target Name/Sectors Targeted', bold)
ws.write('N1', 'Analyst Comments (Specificaly state what the Indicator is being used for[.]  What type of malware/activity it is, and how is it being used within the malware/activity in one or two short paragraphs, Is it a legit but compromised domain?)', bold)
ws.write('O1', 'Additional Notes/Context', bold)

# Prompt users for IOC's that need to be searched.
while inquiring:
    userCommand = input("Please enter an unfanged IOC name to add it to your file, or type 'EXIT' to exit: ")
    if userCommand == "EXIT":
        inquiring = False
    else:
        try:
            methods.publish(userCommand, ws, currentIOC)
            currentIOC += 1
        except:
            print("An API related error has occurred. Please check your input, consider your API limits, and try again.")


workbook.close()