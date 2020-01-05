import numpy as np
from tkinter import *
import datetime
from tkinter import messagebox
import calendar
from dateutil.relativedelta import *
import threading
from tkinter.ttk import Progressbar
import tkinter
import psycopg2 as pg
import os
import csv
import shutil
from PIL import Image, ImageTk
import openpyxl as xl


class FileManager:
    # arranges all interaction with OS files and directories
    def __init__(self, window):
        self.error = Error(window)
        self.get_program_dir()
        self.create_download_dir_name("factuur_exports")
        self.make_dir()

    def get_program_dir(self):
        self.downloads = os.path.expanduser('~\Downloads')
        self.origin_dir = os.path.dirname(os.path.realpath(__file__))

    def get_download_dirs(self):
        try:
            self.directories = []
            for r, d, f in os.walk(self.downloads):
                for directory in d:
                    self.directories.append(directory)
        except:
            self.error.create_window("Fout bij het lezen van de download directories")

    def create_download_dir_name(self, name):
        try:
            i = 1
            new_name = name
            while True:
                unique = True
                self.get_download_dirs()
                for dir in self.directories:
                    if new_name == dir:
                        unique = False
                if not unique:
                    if new_name == name:
                        new_name = name + "(" + str(i) + ")"
                    else:
                        end = new_name[len(name):]
                        number = end[1:-1]
                        new_number = int(number) + 1
                        new_name = name + "(" + str(new_number) + ")"
                    i += 1
                elif unique:
                    self.unique_directory = new_name
                    break
        except:
            self.error.create_window("Fout bij het opstellen van een directory naam")

    def make_dir(self):
        try:
            self.create_download_dir_name("factuur_exports")
            self.new_dir = self.downloads + "\\" + self.unique_directory
            os.mkdir(self.new_dir)
        except:
            self.error.create_window("Probleem met opstellen directory")

    def copy_xls(self, new_name):
        try:
            old_loc = self.origin_dir + "\\resources\sjabloon.xlsx"
            new_loc = self.new_dir + "\\" + new_name + ".xlsx"
            shutil.copy(old_loc, new_loc)
        except:
            self.error.create_window("probleem met kopiëren XLS sjabloon file")


class Date:
    # simple class that holds date information
    def __init__(self, raw_date):
        self.raw_date = raw_date
        self.make_date()

    @staticmethod
    def check_date(raw_date):
        try:
            datetime.datetime.strptime(raw_date, "%d-%m-%Y")
            return True
        except:
            return False

    def make_date(self):
        self.date = datetime.datetime.strptime(self.raw_date, "%d-%m-%Y")


class Calender:
    # creates a calender button group in a predefined frame for date choosing
    def __init__(self, frame, name, entry):
        self.frame = frame
        self.entry = entry
        self.name = name
        self.now = datetime.datetime.now()
        self.selected_period = self.now
        self.make_Frame()

    def make_Frame(self):

        row = 0
        self.name_label = Label(self.frame, text=self.name)
        self.name_label.grid(row=row, column=3)

        row = 1
        self.button_prev_month = Button(self.frame, text="<<<", command=self.previous)
        self.button_prev_month.grid(row=row, column=2)
        self.button_next_month = Button(self.frame, text=">>>", command=self.next)
        self.button_next_month.grid(row=row, column=4)
        self.label_moment = Label(self.frame, text="MM YYYY")
        self.label_moment.grid(row=row, column=3)

        row = 2
        days = ["maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag", "zondag"]
        days_labels = [0 for x in range(len(days))]
        for i in range(len(days)):
            days_labels[i] = Label(self.frame, text=days[i])
            days_labels[i].grid(row=row, column=i)
            days_labels[i].config(width=8)

        row = 3
        self.day_buttons = [0 for x in range(42)]
        for i in range(len(self.day_buttons)):
            self.day_buttons[i] = Button(self.frame, text="")
            self.day_buttons[i].grid(row=row + int(i/7), column=i%7)
            self.day_buttons[i].config(height=2, width=8)
            self.day_buttons[i].bind("<1>", self.chose_date)

        self.label_buttons()

    def label_buttons(self):
        self.label_moment.config(text=str(self.selected_period.year) + " " + self.selected_period.strftime("%b"))
        month = self.selected_period.month
        year = self.selected_period.year
        this_month = datetime.datetime(year, month, 1)
        for i in range(len(self.day_buttons)):
            self.day_buttons[i].config(text="")
        for i in range(calendar.monthrange(year, month)[1]):
           self.day_buttons[i + this_month.weekday()].config(text=str(i + 1))

    def next(self):
        self.selected_period += relativedelta(months=1)
        self.label_buttons()

    def previous(self):
        self.selected_period += -relativedelta(months=1)
        self.label_buttons()

    def chose_date(self, event):
        if event.widget['text'] != "":
            self.chosen_date = datetime.datetime(self.selected_period.year, self.selected_period.month, int(event.widget['text']))
            self.entry.delete(0, "end")
            self.entry.insert(END, str(self.chosen_date.day) + "-" + str(self.chosen_date.month) + "-" + str(self.chosen_date.year))
            self.frame.pack_forget()


class DatePicker:
    # window with date entries and checks, implements calender for choosing options
    def __init__(self):
        self.now = datetime.datetime.now()
        self.selected_period = self.now
        self.chosen_date = "init"
        self.start_date = ""
        self.end_date = ""
        self.make_gui()

    def make_gui(self):
        self.main_window = Tk()
        self.main_window.title("factuur exporteer tool")
        self.main_window.geometry("750x750")
        self.close_window = CloseWindow(self.main_window)
        self.main_window.protocol("WM_DELETE_WINDOW", self.close_window.create_window)

        try:
            # background
            self.load = Image.open(os.path.dirname(os.path.realpath(__file__)) + "\\resources\image1.png")
            self.load = self.load.resize((750, 750), Image.ANTIALIAS)
            self.render = ImageTk.PhotoImage(self.load)
            self.background_label = Label(self.main_window, image=self.render)
            self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

            # frame 0
            self.frame_0 = Frame(self.main_window)
            load = Image.open(os.path.dirname(os.path.realpath(__file__)) + "\\resources\image2.png")
            load = load.resize((int(456/2), int(100/2)), Image.ANTIALIAS)
            render = ImageTk.PhotoImage(load)
            img = Label(self.frame_0, image=render)
            img.image = render
            img.place(x=0, y=0)
            img.pack()
            self.frame_0.pack()
        except:
            pass

        # frame 1
        self.frame_1 = Frame(self.main_window)
        Label(self.frame_1, text="van").grid(row=0, column=0, sticky='W')
        self.start_entry = Entry(self.frame_1)
        self.start_entry.insert(END, "dd-mm-yyyy")
        self.start_entry.bind("<1>", self.activate_1)
        self.start_entry.grid(row=0, column=1, sticky='E')
        Label(self.frame_1, text="tot en met").grid(row=1, column=0)
        self.end_entry = Entry(self.frame_1)
        self.end_entry.insert(END, "dd-mm-yyyy")
        self.end_entry.bind("<1>", self.activate_2)
        self.end_entry.grid(row=1, column=1)
        self.frame_1.pack()

        # frame 2
        self.frame_2 = Frame(self.main_window)
        self.start_calender = Calender(self.frame_2, "start datum", self.start_entry)

        # frame 4
        self.frame_4 = Frame(self.main_window)
        self.end_calender = Calender(self.frame_4, "eind datum", self.end_entry)

        # frame 5
        self.frame_5 = Frame(self.main_window)
        self.start_button = Button(self.frame_5, text="Volgende", command=self.button_action)
        self.start_button.pack()
        self.frame_5.pack()

    def button_action(self):
        self.check_input()
        if self.date_ok == True:
            self.main_window.destroy()

    def check_input(self):
        self.date_ok = False

        # check if input can be transformed to valid date format

        try:
               start_correct = Date.check_date(self.start_entry.get())
        except:
            start_correct = False

        try:
            end_correct = Date.check_date(self.end_entry.get())
        except:
            end_correct = False

        # depending on validity of input continue

        if start_correct is False and end_correct is True:
            messagebox.showerror("Foute datum!", "start datum niet geldig")

        elif start_correct is True and end_correct is False:
            messagebox.showerror("Foute datum", "eind datum niet geldig")

        elif start_correct is False and end_correct is False:
            messagebox.showerror("Foute datum", "begin en eind datum niet geldig")

        elif start_correct and end_correct:
            self.start_date = Date(self.start_entry.get())
            self.end_date = Date(self.end_entry.get())

            if self.start_date.date >= self.end_date.date:
                messagebox.showerror("Ongeldig bereik", "Eind datum moet later zijn dan begin datum")

            else:
                self.date_ok = True

    def activate_1(self, event):
        self.clear_entry(event)
        self.show_calender(event, self.frame_2)
        self.frame_4.pack_forget()

    def activate_2(self, event):
        self.clear_entry(event)
        self.show_calender(event, self.frame_4)
        self.frame_2.pack_forget()

    def clear_entry(self, event):
        if event.widget.get() == "dd-mm-yyyy":
            event.widget.delete(0, "end")

    def show_calender(self, event, frame):
        frame.pack()


class CitySelector:
    # window for choosing cities
    def __init__(self, city_array):
        self.cities = city_array
        self.make_gui()
        self.selection = []

    def make_gui(self):
        self.main_window = Tk()
        self.main_window.title("factuur exporteer tool")
        self.main_window.geometry("750x750")
        self.close_window = CloseWindow(self.main_window)
        self.main_window.protocol("WM_DELETE_WINDOW", self.close_window.create_window)

        try:
            # background
            self.load = Image.open(os.path.dirname(os.path.realpath(__file__)) + "\\resources\image.png")
            self.load = self.load.resize((750, 750), Image.ANTIALIAS)
            self.render = ImageTk.PhotoImage(self.load)
            self.background_label = Label(self.main_window, image=self.render)
            self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

            # frame 0
            self.frame_0 = Frame(self.main_window)
            load = Image.open(os.path.dirname(os.path.realpath(__file__)) + "\\resources\Geosparc.png")
            load = load.resize((int(456/2), int(100/2)), Image.ANTIALIAS)
            render = ImageTk.PhotoImage(load)
            img = Label(self.frame_0, image=render)
            img.image = render
            img.place(x=0, y=0)
            img.pack()
            self.frame_0.pack()
        except:
            pass

        self.frame_1 = Frame(self.main_window)
        self.label = Label(self.frame_1, text="Kies uit de volgende gemeenten:")
        self.label.grid(row=0, sticky='W')
        self.checkboxes = []
        self.values = []
        self.frame_1.pack()

        self.frame_2 = Frame(self.main_window)
        for i in range(len(self.cities)):
            var = IntVar()
            self.checkboxes.append(Checkbutton(self.frame_2, text=self.cities[i, 0], variable=var))
            self.values.append(var)

        for i in range(len(self.checkboxes)):
            self.checkboxes[i].grid(row=int(i/5)+1, column=(i % 5), sticky='W')

        self.frame_2.pack()

        self.frame_3 = Frame(self.main_window)
        self.button = Button(self.frame_3, text="Volgende", command=self.next)
        self.button.grid(sticky=S)
        self.warning = Label(self.frame_3, text="Geen gemeenten geselecteerd!")
        self.frame_3.pack()

    def next(self):
        for i in range(len(self.values)):
            if self.values[i].get() == 1:
                self.selection.append(self.cities[i])
        if len(self.selection) == 0:
            if not self.warning.winfo_ismapped():
                self.warning.grid(sticky=S)
        else:
            self.selection = np.array(self.selection)
            self.main_window.destroy()


class Login:
    # used to login in database, implements progressbar class while validating
    def __init__(self, selection):
        self.host = "host"
        self.databases = selection[:, 1]
        self.port = "portnumber"
        self.successful = False
        self.connections = []
        self.make_gui()
        self.error = Error(self.main_window)

    def make_gui(self):
        self.valid = False
        self.main_window = Tk()
        self.main_window.title("factuur exporteer tool")
        self.main_window.geometry("750x750")

        try:
            self.load = Image.open(os.path.dirname(os.path.realpath(__file__)) + "\\resources\Spotbooking.png")
            self.load = self.load.resize((750, 750), Image.ANTIALIAS)
            self.render = ImageTk.PhotoImage(self.load)
            self.background_label = Label(self.main_window, image=self.render)
            self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

            # frame 0
            self.frame_0 = Frame(self.main_window)
            load = Image.open(os.path.dirname(os.path.realpath(__file__)) + "\\resources\Geosparc.png")
            load = load.resize((int(456/2), int(100/2)), Image.ANTIALIAS)
            render = ImageTk.PhotoImage(load)
            img = Label(self.frame_0, image=render)
            img.image = render
            img.place(x=0, y=0)
            img.pack()
            self.frame_0.pack()
        except:
            pass

        # frame 1
        self.frame_1 = Frame(self.main_window)
        self.user_label = Label(self.frame_1, text='Gebruikersnaam')
        self.user_entry = Entry(self.frame_1)
        self.user_entry.insert(END, "login")
        self.user_entry.bind("<1>", self.clear_entry)
        self.pass_label = Label(self.frame_1, text='Wachtwoord')
        self.pass_entry = Entry(self.frame_1, show='*', text='Wachtwoord')
        self.pass_entry.insert(END, "password")
        self.pass_entry.bind("<1>", self.clear_entry)
        self.val_button = Button(self.frame_1, text='valideren', command=self.validate)

        self.user_label.grid(row=0, column=0)
        self.user_entry.grid(row=0, column=1)
        self.pass_label.grid(row=1, column=0)
        self.pass_entry.grid(row=1, column=1)
        self.val_button.grid(row=2, column=1)

        self.no_connect = Label(self.frame_1, text='kon geen connectie maken met database...')

        self.frame_1.pack()

        # frame 2
        self.frame_2 = Frame(self.main_window)

        # frame 3
        self.frame_3 = Frame(self.main_window)
        label = Label(self.frame_3, text="Connectie met database is succesvol")
        label.pack()
        next_button = Button(self.frame_3, text='Volgende', command=self.next)
        next_button.pack()

        # exit options
        self.close_window = CloseWindow(self.main_window)
        self.main_window.protocol("WM_DELETE_WINDOW", self.close_window.create_window)

    def clear_entry(self, event):
        if event.widget.get() == "Wachtwoord" or event.widget.get() == "Gebruikersnaam":
            event.widget.delete(0, "end")

    def next(self):
        self.main_window.destroy()

    def validate(self):
        try:
            # get user input
            self.user = self.user_entry.get()
            self.pas = self.pass_entry.get()

            # adapt window
            self.frame_1.pack_forget()

            # create processing object
            for widget in self.frame_2.winfo_children():
                widget.destroy()
            self.processing = ProcessBar(self.connect_db, self.frame_2, "test connectie met database")
            self.frame_2.pack()
            self.processing.make_frame()
            self.processing.workflow()
        except:
            self.error.create_window("Fout bij het valideren van de login gegevens")

    def connect_db(self):
        try:
            for i in range(len(self.databases)):
                self.pgconn = pg.connect("dbname={} host={} port={} user={} password={}".format(self.databases[i], self.host, self.port, self.user, self.pas))
                self.connections.append(self.pgconn)
            self.successful = True
            self.frame_2.pack_forget()
            self.frame_3.pack()

        except:
            self.successful = False
            self.frame_2.pack_forget()
            self.frame_1.pack()
            if not self.no_connect.winfo_ismapped():
                self.no_connect.grid(row=4, column=1)


class ProcessBar:
    # creates a progressbar widget in a thread that runs simultaneous with a predefined other task in a second threat
    def __init__(self, task, frame, label):
        self.label = label
        self.task = task
        self.frame = frame

    def make_frame(self):
        self.label = Label(self.frame, text=self.label)
        self.label.pack()
        self.prg_bar = Progressbar(self.frame, orient='horizontal', mode='indeterminate')
        self.prg_bar.pack()
        self.prg_bar.start(15)

        self.t2 = threading.Thread(target=self.task)
        self.t2.setDaemon(True)   # make the thread close when sys.exit is called
        self.t1 = threading.Thread(target=self.progressbar)
        self.t1.setDaemon(True)

    def progressbar(self):
        while True:
            if not self.t2.is_alive():
                self.prg_bar.stop()
                break

    def workflow(self):
        self.t2.start()
        self.t1.start()


class Query:
    # connect to database and execute query function
    def __init__(self, connections, start_date, end_date):
        self.connections = connections
        self.start_date = start_date
        self.end_date = end_date
        self.exports = []

    def execute(self):
        for i in range(len(self.connections)):
            statement = """select *""".format(self.start_date, self.end_date)

            c = self.connections[i].cursor()
            c.execute(statement)
            data = c.fetchall()
            self.exports.append(data)

            c.close()
            self.connections[i].commit()
            self.connections[i].close()


class Export:
    # window that shows up when exporting the data to csv of .xlsx files
    # runs also in thread simultaneously with progress bar widget
    def __init__(self, login_connections, selection, start_date, end_date, mode):
        self.mode = mode
        self.query = Query(login_connections, date_picker.start_date.date, date_picker.end_date.date)
        self.selection = selection
        self.start_date = start_date
        self.end_date = end_date
        self.make_gui()
        self.file_manager = FileManager(self.main_window)
        self.error = Error(self.main_window)
        self.do()

    def make_gui(self):
        self.main_window = Tk()
        self.main_window.title("factuur exporteer tool")
        self.main_window.geometry("750x750")

        try:
            self.load = Image.open(os.path.dirname(os.path.realpath(__file__)) + "\\resources\Spotbooking.png")
            self.load = self.load.resize((750, 750), Image.ANTIALIAS)
            self.render = ImageTk.PhotoImage(self.load)
            self.background_label = Label(self.main_window, image=self.render)
            self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

            # frame 0
            self.frame_0 = Frame(self.main_window)
            load = Image.open(os.path.dirname(os.path.realpath(__file__)) + "\\resources\Geosparc.png")
            load = load.resize((int(456/2), int(100/2)), Image.ANTIALIAS)
            render = ImageTk.PhotoImage(load)
            img = Label(self.frame_0, image=render)
            img.image = render
            img.place(x=0, y=0)
            img.pack()
            self.frame_0.pack()
        except:
            pass

        self.frame_1 = Frame(self.main_window)
        self.frame_1.pack()

        self.frame_2 = Frame(self.main_window)
        self.label = Label(self.frame_2, text='Succesvol bestand geschreven')
        self.label.pack()
        self.button = Button(self.frame_2, text='Afsluiten', command=self.close)
        self.button.pack()
        self.close = CloseWindow(self.main_window)
        self.main_window.protocol("WM_DELETE_WINDOW", self.close.create_window)

    def close(self):
        self.main_window.destroy()
        sys.exit()

    def write_csv(self):
        try:
            self.query.execute()
        except:
            self.error.create_window("Fout bij uitvoeren query")
        try:
            for i in range(len(self.query.exports)):
                name = str(self.selection[i, 0]) + "_" + str(self.start_date.year) + "-" + str(self.start_date.month) + "-" \
                       + str(self.start_date.day) + "_tot_" + str(self.end_date.year) + "-" + str(self.end_date.month) + \
                       "-" + str(self.end_date.day) + ".csv"
                with open(self.file_manager.new_dir + "\\" + name, mode='w', newline='') as csv_file:
                    csv_writer = csv.writer(csv_file, delimiter=";")
                    for row in range(len(self.query.exports[i])):
                        csv_writer.writerow(self.query.exports[i][row])
            self.frame_1.pack_forget()
            self.frame_2.pack()

        except:
            self.error.create_window("Fout bij het schrijven van csv bestand")

    def write_xls(self):
        try:
            self.query.execute()
        except:
            self.error.create_window("Fout bij uitvoeren van query")

        #try:
        self.to_many =[]
        for i in range(len(self.query.exports)):
            name = str(self.selection[i, 0]) + "_" + str(self.start_date.year) + "-" + str(self.start_date.month) \
                   + "-" + str(self.start_date.day) + "_tot_" + str(self.end_date.year) + "-" + \
                   str(self.end_date.month) + "-" + str(self.end_date.day)

            wb = xl.load_workbook(self.file_manager.origin_dir + "\\resources\\sjabloon.xlsx")
            ws = wb["Detail gebruik SB"]

            for row in range(len(self.query.exports[i])):
                for element in range(len(self.query.exports[i][row])):
                    cell = ws.cell(row=row + 16, column=element + 1)
                    cell.value = self.query.exports[i][row][element]

            wb.save(self.file_manager.new_dir + "\\" + name + ".xlsx")

            if len(self.query.exports[i]) >= 1:
                self.to_many.append(self.selection[i][0])

        self.frame_1.pack_forget()
        if len(self.to_many) > 2000:
                messagebox.showwarning("OPGELET", "In volgende bestanden moet het berijk van de formules nog manueel uitgebreid worden: \n-" + " \n-".join(self.to_many))

        self.frame_2.pack()
        #except:
            #self.error.create_window("Fout bij het schrijven naar xlsx bestand ")

    def do(self):
        try:
            if self.mode == 1:
                self.proc = ProcessBar(self.write_csv, self.frame_1, "Data ophalen en schrijven naar csv")
            elif self.mode == 2:
                self.proc = ProcessBar(self.write_xls, self.frame_1, "Data ophalen en schrijven naar xlsx")
            else:
                self.error.create_window("Foute optie keuze CSV of XLS moet 1 of 2 zijn")
            self.proc.make_frame()
            self.proc.workflow()
        except:
            self.error.create_window("Algemene fout bij het uitvoeren van de data verwerking")


class CloseWindow:
    # popup window that can be used for main windows when the default close option is clicked
    def __init__(self, main_window):
        self.main_window = main_window

    def create_window(self):
        self.popup = tkinter.Toplevel(self.main_window)
        self.popup.title("Exit?")
        self.popup.geometry('250x100')
        label = Label(self.popup, text='Programma afsluiten?')
        label.grid(row=1, column=1)
        close_button = Button(self.popup, text='Afsluiten', command=self.close)
        close_button.grid(row=2, column=1)
        cancel_button = Button(self.popup, text='Annuleren', command=self.cancel)
        cancel_button.grid(row=2, column=2)
        self.popup.protocol("WM_DELETE_WINDOW", self.cancel)

    def close(self):
        sys.exit()

    def cancel(self):
        self.popup.destroy()


class Error:
    # popup window when an unforeseen error occurs, closes the whole program
    def __init__(self, main_window):
        self.main_window = main_window

    def create_window(self, message):
        self.popup = tkinter.Toplevel(self.main_window)
        self.popup.title("  ¯\_(ツ)_/¯  ERROR")
        self.popup.geometry('400x200')
        self.label = Label(self.popup, text="Er is iets mis gegaan")
        self.label.pack()
        self.message_label = Label(self.popup, text="foutmelding: " + message)
        self.message_label.pack()
        self.extra = Label(self.popup, text="Man Man Man, miserie, miserie miserie...")
        self.extra.pack()
        self.exit_button = Button(self.popup, text="Afsluiten", command=self.exit)
        self.exit_button.pack()
        self.close = CloseWindow(self.popup)
        self.popup.protocol("WM_DELETE_WINDOW", self.close.create_window)

    def exit(self):
        self.popup.destroy()
        self.main_window.destroy()
        sys.exit()


# date selection
date_picker = DatePicker()
date_picker.main_window.mainloop()

# city selection
# this can be adapted, array of type [["name", "database"], ["name2", "database2"] ...]
cities = np.array([
    ["Atlantis", "db"],
    ["Neverland", "db"],
    ["Smurfendorp", "db"],
    ["Sin city", "db"],
    ["El Dorado", "db"],
    ["Gotham", "db"],
    ["Duckburg", "db"],
    ["Wonderland", "db"],
    ["Camelot", "db"]
    ])

city_selector = CitySelector(cities)
city_selector.main_window.mainloop()

# login with username and password
login = Login(city_selector.selection)
login.main_window.mainloop()

# execute
# mode = 1: write csv, mode = 2: write xls
export = Export(login.connections, city_selector.selection, date_picker.start_date.date, date_picker.end_date.date, mode=2)
export.main_window.mainloop()


def test_write(mode):
    start = datetime.date(2019, 1, 1)
    end = datetime.date(2019, 12, 31)

    cities = np.array([
        ["Atlantis", "db"],
        ["Troje", "db"]
        ])
    connections = []

    for i in range(len(cities)):
        pgconn = pg.connect("dbname={} host={} port={} user={} password={}".format(cities[i, 1], "db", "port", "login", "password"))
        connections.append(pgconn)

    query = Query(connections, start, end)
    export = Export(query, cities, start, end, mode)
    export.main_window.mainloop()

