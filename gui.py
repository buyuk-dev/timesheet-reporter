import tkinter
import sysutils

import send_mail_outlook
import send_mail
import writer
import common
import model

import config


class RootFrame(tkinter.Frame):
    def __init__(self, parent, *args, **kwargs):
        tkinter.Frame.__init__(self, parent.tk_context, *args, **kwargs)
        self.parent = parent

        
class AppMenu(tkinter.Frame):

    def __init__(self, parent, *args, **kwargs):
        tkinter.Frame.__init__(self, parent, *args, **kwargs)

        self.parent = parent
        app_instance = self.parent.parent

        self.save = tkinter.Button(self, text=config.save_btn_label)
        self.save.configure(command = lambda : app_instance.onSave())
        self.save.pack(side='left')

        self.load = tkinter.Button(self, text=config.load_btn_label)
        self.load.configure(command = lambda : app_instance.onLoad())
        self.load.pack(side='left')

        self.send = tkinter.Button(self, text=config.send_btn_label)
        self.send.configure(command = lambda : app_instance.onSend())
        self.send.pack(side='left')

        self.addReceipient = tkinter.Button(self, text=config.add_receipient_btn_label)
        self.addReceipient.configure(command = lambda: app_instance.onNewReceipient())
        self.addReceipient.pack(side='left')

        self.generate = tkinter.Button(self, text=config.generate_btn_label)
        self.generate.configure(command = lambda : app_instance.onGenerate())
        self.generate.pack(side='left')
        
        self.reset = tkinter.Button(self, text=config.reset_btn_label)
        self.reset.configure(command = lambda: app_instance.onReset())
        self.reset.pack(side='left')

        self.pack()        

        
class GuiApp:

    def __init__(self):
        self.tk_context = tkinter.Tk()
        self.tk_context.title(config.app_title)
        self.tk_context.protocol("WM_DELETE_WINDOW", self.tk_context.quit)
        self.onReset()

    def initEntries(self, entries=None):
        self.entriesHeader = tkinter.Label(self.rootFrame, text=config.entries_section_label)
        self.entriesHeader.pack()
        self.entries = list()
        if entries is None:
            for date in common.get_dates_range(config.monday, config.week_length):
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
        try:
            self.rootFrame.destroy()
        except:
            pass

        self.rootFrame = RootFrame(self)

        try:
            self.dataModel.reset()
        except:
            self.dataModel = model.DataModel()

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
        self.menu = AppMenu(self.rootFrame)
        self.initReceipients(data["receipients"])
        self.initEntries(data["entries"])

    def onSend(self):
        if config.use_outlook:
            send_mail_outlook.send_mail(
                config.sender_address,
                self.dataModel.getReceipients(),
                config.message_title,
                config.message_body,
                [config.xls_path]
            )
        else:
            send_mail.send_mail(
                config.sender_address,
                self.dataModel.getReceipients(),
                config.message_title,
                config.message_body,
                [config.xls_path],
                config.credentials
            )

    def onReset(self):
        self.reset()
        self.menu = AppMenu(self.rootFrame)
        self.initReceipients()
        self.initEntries()

    def run(self):
        tkinter.mainloop()
