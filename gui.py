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
        self.pack()

        
class AppMenu(tkinter.Frame):

    def __init__(self, parent, menupady=5, btnpadx=5, *args, **kwargs):
        tkinter.Frame.__init__(self, parent, pady=menupady, bg="grey", *args, **kwargs)

        self.parent = parent
        app_instance = self.parent.parent

        self.save = tkinter.Button(self, text=config.save_btn_label, bg="orange")
        self.save.configure(command = lambda : app_instance.onSave())
        self.save.pack(side='left', padx=btnpadx)

        self.load = tkinter.Button(self, text=config.load_btn_label)
        self.load.configure(command = lambda : app_instance.onLoad())
        self.load.pack(side='left', padx=btnpadx)

        self.send = tkinter.Button(self, text=config.send_btn_label)
        self.send.configure(command = lambda : app_instance.onSend())
        self.send.pack(side='left', padx=btnpadx)

        self.addReceipient = tkinter.Button(self, text=config.add_receipient_btn_label)
        self.addReceipient.configure(command = lambda: app_instance.receipients.add_receipient())
        self.addReceipient.pack(side='left', padx=btnpadx)

        self.generate = tkinter.Button(self, text=config.generate_btn_label)
        self.generate.configure(command = lambda : app_instance.onGenerate())
        self.generate.pack(side='left', padx=btnpadx)
        
        self.reset = tkinter.Button(self, text=config.reset_btn_label)
        self.reset.configure(command = lambda: app_instance.onReset(), bg="red", fg="white")
        self.reset.pack(side='left', padx=btnpadx)

        self.pack()        


class WorkEntries(tkinter.Frame):

    def __init__(self, parent, entries=None, *args, **kwargs):
        tkinter.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.pack()
        
        self.entriesHeader = tkinter.Label(self, text=config.entries_section_label)
        self.entriesHeader.pack()

        self.entries = list()

        if entries is None:
            for date in common.get_dates_range(config.monday, config.week_length):
                self.add_entry(date)
        else:
            for date, entry in entries.items():
                self.add_entry(date, entry["hours"], entry["description"])

    def add_entry(self, date, hour=8, description=""):
        print("adding entry: {}".format((date, hour, description)))

        app_instance = self.parent.parent

        var_entry, var_hours = app_instance.dataModel.addEntry(date, description, hour)

        frame = tkinter.Frame(self, width=config.entries_frame_width)
        frame.pack()

        label = tkinter.Label(frame, text=date.strftime(config.date_label_format), width=config.date_label_width, anchor='w')
        label.pack(side="left")

        hour = tkinter.Entry(frame, width=config.hours_entry_width, textvariable=var_hours)
        hour.pack(side="left")

        entry = tkinter.Entry(frame, width=config.description_entry_width, textvariable=var_entry)
        entry.pack(side="left")

        self.entries.append((label, hour, entry))

        
class Receipients(tkinter.Frame):

    def __init__(self, parent, receipients=[], *args, **kwargs):
        tkinter.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.pack()

        self.recpLabel = tkinter.Label(self, text=config.receipients_section_label)
        self.recpLabel.pack()

        self.receipientEntries = []
        for recp in receipients:
            self.add_receipient(recp)

    def add_receipient(self, receipient=""):
        app_instance = self.parent.parent
        var = app_instance.dataModel.addReceipient(receipient)
        entry = tkinter.Entry(self, width=config.receipient_entry_width, textvariable=var)
        self.receipientEntries.append(entry)
        entry.pack()
        if receipient is not None:
            var.set(receipient)


class GuiApp:

    def __init__(self):
        self.tk_context = tkinter.Tk()
        self.tk_context.title(config.app_title)
        self.tk_context.protocol("WM_DELETE_WINDOW", self.tk_context.quit)
        self.dataModel = model.DataModel()
        self.initGui()

    def initGui(self, receipients=config.receipients, entries=None):
        self.rootFrame = RootFrame(self)
        self.menu = AppMenu(self.rootFrame)
        self.receipients = Receipients(self.rootFrame, receipients)
        self.work_entries = WorkEntries(self.rootFrame, entries)
        
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
        self.rootFrame.destroy()
        self.dataModel.reset()
        data = self.dataModel.load(config.save_path)
        self.initGui(data["receipients"], data["entries"])

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
        self.rootFrame.destroy()
        self.dataModel.reset()
        self.initGui()

    def run(self):
        tkinter.mainloop()
