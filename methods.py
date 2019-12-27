# methods.py
#
# A part of the IOC Assistant software package, developed by Brandon Miller (DHS-CISA Intern 2019-2020)
from __future__ import print_function
import json
import whois
import socket
from riskiq.api import Client
import datetime
import xlsxwriter
import hashlib
from virus_total_apis import PublicApi as VirusTotalPublicApi


def publish(IOC, sheet, num):
    # Make preliminary API calls.
    vtdump = getVTDump(IOC)
    whoisdump = getWhoIsDump(IOC)



    # A - Publish Indicator
    cell = "A" + str(num)
    sheet.write(cell, IOC)

    # B - Here, we skip IOC Source Report Name and Report Location for the analyst to assess.

    # C - Publish Activity Type
    # TODO

    # D - Here, we skip malware name for the analyst to assess.

    # E - Here, we skip Indicator Type for the analyst to assess.

    # F - Here, we skip Actor Name for the analyst to assess.

    # G - Associated IP address
    #IP = ""
    #try:
    #    IP = socket.gethostbyname(IOC)
    #    cell = "G" + str(num)
    #    sheet.write(cell, "Currently resolves to " + IP)
    #except:
    #    cell = "G" + str(num)
    #    sheet.write(cell, "Does not currently resolve.")

    # H - Virus Total Ratio
    cell = "H" + str(num)
    sheet.write(cell, getRatio(vtdump))

    # I - Alexa Page Rank
    cell = "I" + str(num)

    # J - Total Number of Subdomains
    # TODO

    # K - Domain Registrant (Owner)
    # TODO

    # L - Domain Registration and Expiration Date
    cell = "L" + str(num)
    sheet.write(cell, getWhoIsDates(whoisdump))

    # M - Here, we skip Target Name/Sectors Targeted for the analyst to assess.

    # N - Here, we skip Analyst Comments for obvious reasons.

    # O - Here, we skip additional notes/context, for obvious reasons.


def getVTDump(IOC):
    # Prepare Virus Total API for use.
    key = 'Insert VT Api Key Here'
    vt = VirusTotalPublicApi(key)
    response = vt.get_url_report(IOC)

    # Get and return dump.
    dump = json.dumps(response, sort_keys=False, indent=4)
    return dump


def getWhoIsDump(IOC):
    domain = whois.whois(IOC)
    return domain


def getRatio(dump):
    y = json.loads(dump)
    # Derive necessary information from json dump.
    try:
        ratio = str(y["results"]["positives"]) + "/" + str(y["results"]["total"])
        return ratio
    except:
        return "N/A"


def getWhoIsDates(dump):
    s = "Created on " + str(dump.creation_date) + " and expires on " + str(dump.expiration_date)
    return s
