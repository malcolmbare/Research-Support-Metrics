from argparse import RawDescriptionHelpFormatter
from asyncio import SendfileNotAvailableError
from datetime import datetime
from dateutil.relativedelta import relativedelta
from msal import PublicClientApplication
from requests_oauthlib import OAuth2Session
from oauthlib.oauth2 import BackendApplicationClient
import requests
import csv
from ReqAndAuth import outlook
from collections import Counter
import regex as re
import os
import sys

class conversationRange:
    def addMonth(self, startDate):
        dto = datetime.fromisoformat(startDate[:-1])
        dto += relativedelta(months=1)
        return dto.isoformat() + "Z"
    def getConversationIDs(self, startDate, endDate):
        archiveEndpoint = self.baseEndpoint + "messages?$filter=sentdatetime ge {0} and sentdatetime lt {1}&top=300".format(startDate,endDate)
        resp = requests.get(archiveEndpoint, headers=self.outlookClient.header).json()
        result = list(set([i["conversationId"] for i in resp["value"]]))
        return result
    def getConversations(self, conversationIDs):
        result = {}
        count = 0
        for cid in conversationIDs:
            conversationEndpoint = self.baseEndpoint + "/messages?filter=conversationID eq '{0}'".format(cid)
            resp = requests.get(conversationEndpoint, headers=self.outlookClient.header).json()
            result[cid] = conversation(resp, self.registry)
            count += 1
            print("Conversation {0} downloaded".format(str(count)))
        return result
    def __init__(self, mm, yyyy, registry):
        self.startDate = "{0}-{1}-01T00:00:00Z".format(yyyy, mm)
        self.endDate = self.addMonth(self.startDate)
        self.outlookClient = outlook()
        self.baseEndpoint = self.outlookClient.serverURL + self.outlookClient.extensions["getArchive"]
        self.baseEndpoint = self.baseEndpoint.format(self.outlookClient.ids["rSupportUser"], self.outlookClient.ids["rSupportArchive"])
        self.conversationIDs = self.getConversationIDs(self.startDate, self.endDate)
        self.registry = registry
        self.conversations = self.getConversations(self.conversationIDs)
        
class conversation:
    def makeMessages(self):
        messages = []
        for msg in self.rawBody:
            currentMessage = message(msg, self.conversationLength)
            messages.append(currentMessage)
        return messages
    def checkRegistry(self, emailAddress):
        if emailAddress in self.registry.registry:
            result = self.registry.registry[emailAddress]
            return result
        else:
            self.registry.notInRegistry.append(emailAddress)
            return "???"
    def __init__(self, conversationDict, registry):
        self.id = conversationDict["value"][0]["conversationId"]
        self.registry = registry
        self.context = conversationDict["@odata.context"]
        self.rawBody = conversationDict["value"]
        self.conversationLength = len(self.rawBody)
        self.messages = self.makeMessages()

class message:
    def indexMessageHistory(self, linedText):
        result = [0]
        for line in range(len(linedText)):
            if linedText[line].startswith("From:") or linedText[line].startswith("wrote:"):
                result.append(line)
        result.append(len(linedText))
        return result
    def unpackAddress(self, recipients):
        result= [(i["emailAddress"]["name"], i["emailAddress"]["address"]) for i in recipients]
        return list(set(result))
    def __init__(self, msg, conversationLength):
        self.id = msg["id"]
        self.conversationId = msg["conversationId"]
        self.conversationLength = conversationLength
        self.receivedTime = msg["receivedDateTime"]
        self.sentTime = msg["sentDateTime"]
        self.subject = msg["subject"]
        self.recipients = self.unpackAddress(msg["toRecipients"] + msg["ccRecipients"])
        self.sender = self.unpackAddress([msg["sender"]] + [msg["from"]])
        self.body= [i for i in msg["body"]["content"].splitlines() if len(i) > 1]

class staffRegistry:
    def openRegistry(self):
        result ={}
        with open("emailMap.csv", newline="") as csvfile:
            cReader = csv.reader(csvfile, delimiter=",")
            for row in cReader:
                result[row[0]] = row[1]
        return(result)
    def __init__(self):
        self.registry = self.openRegistry()
        self.notInRegistry = []

class messageRow:
    def configureDay(self, isoDate):
        dayC = isoDate.split("T")[0]
        components = dayC.split("-")
        return components[1] + "/" + components[2] + "/" + components[0]
    def __init__(self, msg, registry):
        self.registry = registry
        self.conversationId = msg.conversationId
        self.receivedDateTime = msg.receivedTime
        self.receivedDateDay = self.configureDay(self.receivedDateTime)
        self.sentDateTime = msg.sentTime
        self.sentDateDay = self.configureDay(self.sentDateTime)
        self.sender = " | ".join([msg.sender[0][0]])
        self.senderEmail = " | ".join([msg.sender[0][1]])
        self.senderAffiliation = " | ".join([self.registry.registry[i[1]] for i in msg.sender if i[1] in self.registry.registry])
        self.recipients = " | ".join([i[0] for i in msg.recipients])
        self.recipientEmail = " | ".join([i[1] for i in msg.recipients])
        self.recipientAffiliation = " | ".join([self.registry.registry[i[1]] for i in msg.recipients if i[1] in self.registry.registry])
        self.affiliationSimplified = ""
        self.location = ""
        self.issue = ""
        self.package = ""
        self.platform = ""
        self.subject = msg.subject
        self.chainLength = msg.conversationLength
        self.body = " ".join(msg.body)
        self.row = list(self.__dict__.values())[1:] 

class csvExport:
    def exportRegistryRows(self):
        rows = [["email","position"]]
        rows += [[item[0],item[1]] for item in list(self.registry.registry.items())]
        rows += [[email, ""] for email in self.registry.notInRegistry]
        return(rows)
    def exportLabelRows(self):
        rows = [["text","label"]]
        for cKey in self.conversationKeys:
            msg = self.metrics.conversations[cKey].messages[0]
            rows.append(["SUBJ: " + msg.subject +"/n"+ "BODY: " + "/n".join(msg.body), ""])
        return(rows)
    def exportSpreadsheetRows(self):
        rows = []
        for cKey in self.conversationKeys:
            msg = self.metrics.conversations[cKey].messages[0]
            rows.append(messageRow(msg, self.registry).row)
        return(rows)
    def writeCSV(self, month, year, title ,rows):
        path = os.path.join(self.dir,"metrics{0}-{1}-{2}.tsv".format(title,month,year))
        with open(path, "w") as f:
            write = csv.writer(f, delimiter="\t")
            write.writerows(rows)
    def __init__(self, metrics, registry, month, year):
        self.conversationKeys = list(metrics.conversationIDs)
        self.registry = registry
        self.metrics = metrics
        self.registryRows = self.exportRegistryRows()
        self.labelRows = self.exportLabelRows()
        self.spreadsheetRows = self.exportSpreadsheetRows()
        desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
        self.dir =desktop+ "/{0}-{1}-ResearchSupportMetrics".format(month,year)
        os.mkdir(self.dir)
        self.writeCSV(month, year, "EmailMap",self.registryRows)
        self.writeCSV(month,year, "LabelReady",self.labelRows)
        self.writeCSV(month,year, "Spreadsheet",self.spreadsheetRows)

def main():
    registry = staffRegistry()
    if len(sys.argv[1]) ==1:
        month = "0"+ sys.argv[1]
    else:
        month = sys.argv[1]
    year = sys.argv[2]
    metrics = conversationRange(month, year, registry)
    csvExport(metrics, registry, month, year)

if __name__ == "__main__":
    main()


