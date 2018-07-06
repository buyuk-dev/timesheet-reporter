import tkinter
import sysutils
import os

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

        
class AppMenu(tkinter.Menu):

    def __init__(self, parent, menupady=5, btnpadx=5, *args, **kwargs):
        tkinter.Menu.__init__(self, *args, **kwargs)

        parent.tk_context.option_add('*tearOff', tkinter.FALSE)
        self.parent = parent

        # for macos compatibility
        if parent.tk_context.tk.call('tk', 'windowingsystem') == "aqua": 
            menu = tkinter.Menu(self)
            self.add_cascade(label="menu", menu=menu, accelerator="M")
        else:
            menu = self

        menu.add_command(label=config.save_btn_label, command=lambda: parent.onSave(), accelerator="Ctrl-s")
        self.bind_all("<Control-s>", lambda x: parent.onSave())

        menu.add_command(label=config.load_btn_label, command=lambda: parent.onLoad(), accelerator="Ctrl-l")
        self.bind_all("<Control-l>", lambda x: parent.onLoad())

        menu.add_command(label=config.send_btn_label, command=lambda: parent.onSend(), accelerator="Ctrl-m")
        self.bind_all("<Control-m>", lambda x: parent.onSend())

        menu.add_command(label=config.add_receipient_btn_label, command=lambda: parent.receipients.add_receipient(), accelerator="Ctrl-e")
        self.bind_all("<Control-e>", lambda x: parent.receipients.add_receipient())

        menu.add_command(label=config.generate_btn_label, command=lambda: parent.onGenerate(), accelerator="Ctrl-g")
        self.bind_all("<Control-g>", lambda x: parent.onGenerate())

        menu.add_command(label=config.reset_btn_label, command=lambda: parent.onReset(), accelerator="Ctrl-r")
        self.bind_all("<Control-r>", lambda x: parent.onReset())


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
        self.menu = AppMenu(self)
        self.tk_context.config(menu=self.menu)
        self.initGui()

    def initGui(self, receipients=config.receipients, entries=None):
        self.rootFrame = RootFrame(self, padx=10, pady=10)
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
