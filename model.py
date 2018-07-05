from tkinter import StringVar, IntVar
import json
from datetime import datetime, timedelta

import common


class DataModel:

    def __init__(self):
        self.receipients = list()
        self.entries = dict()

    def addReceipient(self, recp_val=""):
        recp = StringVar()
        recp.set(recp_val)
        self.receipients.append(recp)
        return recp

    def getReceipients(self):
        return [recp.get() for recp in self.receipients]

    def addEntry(self, date, desc_val="", hours_val=8):
        entry = StringVar()
        entry.set(desc_val)

        hours = IntVar()
        hours.set(hours_val)

        self.entries[date] = {
            "description": entry,
            "hours": hours
        }
        return entry, hours

    def getEntries(self):
        return {
            date.date(): {
                "description": entry["description"].get(),
                "hours": entry["hours"].get()
            }
            for date, entry in self.entries.items()
        }

    def save(self, filename):
        with open(filename, "w") as json_file:
            data = {
                "receipients": self.getReceipients(),
                "entries": common.convertNonStringKeys(self.getEntries())
            }
            json_str = json.dumps(data, indent=4, sort_keys=True, default=str)
            json_file.write(json_str)

    def load(self, filename):
        with open(filename, "r") as json_file:
            data = json.load(json_file)
            data["entries"] = common.convertBackToDateKeys(data["entries"])
            return data

    def reset(self):
        del self.receipients
        del self.entries
        self.receipients = list()
        self.entries = dict()