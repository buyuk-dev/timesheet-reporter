import tkinter

import config
import sendmail
import writer
import sysutils

from model import DataModel, get_week_dates


class GuiApp:

    def __init__(self):
        self.root = tkinter.Tk()
        self.root.title(config.app_title)
        self.root.protocol("WM_DELETE_WINDOW", self.root.quit)

        self.rootFrame = tkinter.Frame(self.root)
        self.rootFrame.pack()

        self.dataModel = DataModel()

        self.initMenu()
        self.initReceipients()
        self.initEntries()

    def initEntries(self, entries=None):
        self.entriesHeader = tkinter.Label(self.rootFrame, text=config.entries_section_label)
        self.entriesHeader.pack()
        self.entries = list()
        if entries is None:
            for date in get_week_dates():
                self.addEntry(date)
        else:
            for date, entry in entries.items():
                self.addEntry(date, entry["hours"], entry["description"])

    def initReceipients(self, receipients=config.receipients):
        self.receipientsFrame = tkinter.Frame(self.rootFrame)
        self.recpLabel = tkinter.Label(self.receipientsFrame, text=config.receipients_section_label)
        self.recpLabel.pack()
        self.receipientEntries = []
        self.receipientsFrame.pack()
        for recp in receipients:
            self.onNewReceipient(recp)

    def initMenu(self):
        self.menu = tkinter.Frame(self.rootFrame)
        self.generateBtn = tkinter.Button(self.menu, text=config.generate_btn_label)
        self.generateBtn.configure(command = lambda : self.onGenerate())

        self.saveBtn = tkinter.Button(self.menu, text=config.save_btn_label)
        self.saveBtn.configure(command = lambda : self.onSave())

        self.loadBtn = tkinter.Button(self.menu, text=config.load_btn_label)
        self.loadBtn.configure(command = lambda : self.onLoad())

        self.sendBtn = tkinter.Button(self.menu, text=config.send_btn_label)
        self.sendBtn.configure(command = lambda : self.onSend())

        self.addReceipientBtn = tkinter.Button(self.menu, text=config.add_receipient_btn_label, command=lambda: self.onNewReceipient())
        self.resetBtn = tkinter.Button(self.menu, text=config.reset_btn_label, command=lambda: self.onReset())

        self.menu.pack()
        self.saveBtn.pack(side='left')
        self.loadBtn.pack(side='left')
        self.generateBtn.pack(side='left')
        self.sendBtn.pack(side='left')
        self.addReceipientBtn.pack(side='left')
        self.resetBtn.pack(side='left')

    def addEntry(self, date, hour=8, description=""):
        var_entry, var_hours = self.dataModel.addEntry(date, description, hour)

        frame = tkinter.Frame(self.rootFrame, width=config.entries_frame_width)
        frame.pack()

        label = tkinter.Label(frame, text=date.strftime(config.date_label_format), width=config.date_label_width, anchor='w')
        label.pack(side="left")

        hour = tkinter.Entry(frame, width=config.hours_entry_width, textvariable=var_hours)
        hour.pack(side="left")

        entry = tkinter.Entry(frame, width=config.description_entry_width, textvariable=var_entry)
        entry.pack(side="left")

        self.entries.append((label, hour, entry))

    def reset(self):
        self.rootFrame.destroy()
        self.rootFrame = tkinter.Frame(self.root)
        self.dataModel.reset()
        self.rootFrame.pack()

    def onNewReceipient(self, receipient=None):
        var = self.dataModel.addReceipient(receipient)
        entry = tkinter.Entry(self.receipientsFrame, width=config.receipient_entry_width, textvariable=var)
        self.receipientEntries.append(entry)
        entry.pack()
        if receipient is not None:
            var.set(receipient)

    def onGenerate(self):
        entries = self.dataModel.getEntries()
        data = [
            (config.employee_name,
            str(key),
            entries[key]["hours"],
            entries[key]["description"])
            for key in sorted(entries)
        ]
        writer.write_raport(data, config.xls_path)

    def onSave(self):
        self.dataModel.save(config.save_path)

    def onLoad(self):
        self.reset()
        data = self.dataModel.load(config.save_path)
        self.initMenu()
        self.initReceipients(data["receipients"])
        self.initEntries(data["entries"])

    def onSend(self):
        sendmail.send_mail(
            config.sender_address,
            self.dataModel.getReceipients(),
            config.message_title,
            config.message_body,
            [config.xls_path]
        )

    def onReset(self):
        self.reset()
        self.initMenu()
        self.initReceipients()
        self.initEntries()

    def run(self):
        tkinter.mainloop()


if __name__ == '__main__':
    try:
        #sysutils.hide_terminal_window()
        app = GuiApp()
        app.run()
    except Exception as e:
        print(e)
        input()
