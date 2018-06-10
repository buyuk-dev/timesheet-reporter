import tkinter
from datetime import datetime, timedelta
import json

import config
#import sendmail
import writer


def generateDates():
    return list((datetime.today() - timedelta(days=4) + timedelta(days=i)).date() for i in range(5))

def convertNonStringKeys(dict_):
    return {str(key): val for key, val in dict_.items()}

def convertBackToDateKeys(dict_):
    return {datetime.strptime(key, "%Y-%m-%d").date(): val for key, val in dict_.items()}


class DataModel:

    def __init__(self):
        self.receipients = list()
        self.entries = dict()

    def addReceipient(self, recp_val=""):
        recp = tkinter.StringVar()
        recp.set(recp_val)
        self.receipients.append(recp)
        return recp

    def getReceipients(self):
        return [recp.get() for recp in self.receipients]

    def addEntry(self, date, desc_val="", hours_val=8):
        entry = tkinter.StringVar()
        entry.set(desc_val)

        hours = tkinter.IntVar()
        hours.set(hours_val)

        self.entries[date] = {
            "description": entry,
            "hours": hours
        }
        return entry, hours

    def getEntries(self):
        return {
            date: {
                "description": entry["description"].get(),
                "hours": entry["hours"].get()
            }
            for date, entry in self.entries.items()
        }

    def save(self, filename):
        with open(filename, "w") as json_file:
            data = {
                "receipients": self.getReceipients(),
                "entries": convertNonStringKeys(self.getEntries())
            }
            json_str = json.dumps(data, indent=4, sort_keys=True, default=str)
            json_file.write(json_str)

    def load(self, filename):
        with open(filename, "r") as json_file:
            data = json.load(json_file)
            data["entries"] = convertBackToDateKeys(data["entries"])
            return data

    def reset(self):
        del self.receipients
        del self.entries
        self.receipients = list()
        self.entries = dict()


class GuiApp:

    def __init__(self):
        self.root = tkinter.Tk()
        self.root.title("Time Report Creator")
        self.root.protocol("WM_DELETE_WINDOW", self.root.quit)

        self.rootFrame = tkinter.Frame(self.root)
        self.rootFrame.pack()

        self.dataModel = DataModel()

        self.initMenu()
        self.initReceipients()
        self.initEntries()

    def initEntries(self, entries=None):
        self.entriesHeader = tkinter.Label(self.rootFrame, text="Work entries")
        self.entriesHeader.pack()

        self.entries = list()

        if entries is None:
            for date in generateDates():
                self.addEntry(date)
        else:
            for date, entry in entries.items():
                self.addEntry(date, entry["hours"], entry["description"])

    def initReceipients(self, receipients=config.receipients):
        self.receipientsFrame = tkinter.Frame(self.rootFrame)
        self.recpLabel = tkinter.Label(self.receipientsFrame, text="Receipients")
        self.recpLabel.pack()
        self.receipientEntries = []
        self.receipientsFrame.pack()
        for recp in receipients:
            self.onNewReceipient(recp)

    def initMenu(self):
        self.menu = tkinter.Frame(self.rootFrame)
        self.generateBtn = tkinter.Button(self.menu, text='generate')
        self.generateBtn.configure(command = lambda : self.onGenerate())

        self.saveBtn = tkinter.Button(self.menu, text='save')
        self.saveBtn.configure(command = lambda : self.onSave())

        self.loadBtn = tkinter.Button(self.menu, text='load')
        self.loadBtn.configure(command = lambda : self.onLoad())

        self.sendBtn = tkinter.Button(self.menu, text='send')
        self.sendBtn.configure(command = lambda : self.onSend())

        self.addReceipientBtn = tkinter.Button(self.menu, text="+ recp", command=lambda: self.onNewReceipient())
        self.resetBtn = tkinter.Button(self.menu, text="reset", command=lambda: self.onReset())

        self.menu.pack()
        self.saveBtn.pack(side='left')
        self.loadBtn.pack(side='left')
        self.generateBtn.pack(side='left')
        self.sendBtn.pack(side='left')
        self.addReceipientBtn.pack(side='left')
        self.resetBtn.pack(side='left')

    def addEntry(self, date, hour=8, description=""):
        var_entry, var_hours = self.dataModel.addEntry(date, description, hour)

        frame = tkinter.Frame(self.rootFrame, width=100)
        frame.pack()

        label = tkinter.Label(frame, text=date.strftime("%d.%m.%Y - %As"), width=20, anchor='w')
        label.pack(side="left")

        hour = tkinter.Entry(frame, width=5, textvariable=var_hours)
        hour.pack(side="left")

        entry = tkinter.Entry(frame, width=50, textvariable=var_entry)
        entry.pack(side="left")

        self.entries.append((label, hour, entry))

    def reset(self):
        self.rootFrame.destroy()
        self.rootFrame = tkinter.Frame(self.root)
        self.dataModel.reset()
        self.rootFrame.pack()

    def onNewReceipient(self, receipient=None):
        var = self.dataModel.addReceipient(receipient)
        entry = tkinter.Entry(self.receipientsFrame, width=50, textvariable=var)
        self.receipientEntries.append(entry)
        entry.pack()
        if receipient is not None:
            var.set(receipient)

    def onGenerate(self):
        entries = self.dataModel.getEntries()
        data = [
            ("Michal Michalski",
            str(key),
            entries[key]["hours"],
            entries[key]["description"])
            for key in sorted(entries)
        ]
        writer.write_raport(data, config.xls_path)

    def onSave(self):
        self.dataModel.save("data.json")

    def onLoad(self):
        self.reset()
        data = self.dataModel.load(config.save_path)
        self.initMenu()
        self.initReceipients(data["receipients"])
        self.initEntries(data["entries"])

    def onSend(self):
        #sendmail.send_mail(
        #    config.sender_address,
        #    self.dataModel.getReceipients(),
        #    config.message_title,
        #    config.message_body,
        #    [config.xls_path]
        #)
        pass

    def onReset(self):
        self.reset()
        self.initMenu()
        self.initReceipients()
        self.initEntries()

    def run(self):
        tkinter.mainloop()


if __name__ == '__main__':
    app = GuiApp()
    app.run()
