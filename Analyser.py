import numpy as np
import matplotlib.pyplot as plt
import pprint
import pandas
import sys
import operator
from datetime import datetime
import time
import copy
import os
import datetime as dt
import csv


class Record:

    def __init__(self, ticket_ID, klant, bedrijf, onderwerp, datum_aangemaakt, laatste_bericht, inbox,
                 verantwoordelijke, status, project, gewerkte_tijd, type01, prioriteit02, bron03, software01,
                 root_cause02, root_cause_detail03, status01, geplande_versie02, feature_group03, storyp_uren04):
        self.ticket_ID = ticket_ID
        self.klant = klant
        self.bedrijf = bedrijf
        self.ontwerp = onderwerp
        self.datum_aangemaakt = datum_aangemaakt
        self.laatste_bericht = laatste_bericht
        self.inbox = inbox
        self.verantwoordelijke = verantwoordelijke
        self.status = status
        self.project = project
        self.gewerkte_tijd = gewerkte_tijd
        self.type01 = type01
        self.prioriteit02 = prioriteit02
        self.bron03 = bron03
        self.software01 = software01
        self.root_cause02 = root_cause02
        self.root_cause_detail = root_cause_detail03
        self.status01 = status01
        self.geplande_versie02 = geplande_versie02
        self.feature_group03 = feature_group03
        self.storyp_uren04 = storyp_uren04

        # derived variables
        self.get_root_cause()
        self.get_month()
        self.get_year()
        self.get_period()
        self.get_end_date()
        self.get_start_date()
        self.compute_repair_time()

    def print_record(self):
        pprint.pprint(self.__dict__)

    def get_month(self):
        self.date = datetime.strptime(self.datum_aangemaakt, '%Y-%m-%d %H:%M:%S')
        self.month_number = self.date.month
        self.number_to_month(self.month_number)

    def number_to_month(self, number):
        switcher = {
            1: "januari",
            2: "februari",
            3: "maart",
            4: "april",
            5: "mei",
            6: "juni",
            7: "juli",
            8: "augustus",
            9: "september",
            10: "oktober",
            11: "november",
            12: "december"
        }
        self.month = switcher.get(number)

    def get_year(self):
        self.year = self.datum_aangemaakt[0:4]

    def get_period(self):
        self.period = str(self.month) + " " + str(self.year)

    def get_root_cause(self):
        if self.root_cause02 == "Bug software" or self.root_cause02 == "Internal service outage" or self.root_cause02 \
                == "Bug configuration":
            self.root_cause = "Geosparc"

        elif self.root_cause02 == "Configuration request" or self.root_cause02 == "Lack of functional knowledge" or \
                self.root_cause02 == "Wrong information from customer" or self.root_cause02 == "Missing Feature":
            self.root_cause = "Customer"

        elif self.root_cause02 == "External service outage":
            self.root_cause = "External"

        elif self.root_cause02 == "Undetermined":
            self.root_cause = "Undetermined"

        elif self.root_cause02 == "":
            self.root_cause = "Not specified"

        else:
            self.root_cause = "cannot recognize type"

    def get_start_date(self):
        if self.datum_aangemaakt != "":
            try:
                self.start_date = datetime.strptime(self.datum_aangemaakt, '%Y-%m-%d %H:%M:%S')
            except:
                Error = FatalError()
                Error.specific_error(4)
        else:
            self.start_date = ""

    def get_end_date(self):
        if self.laatste_bericht != "":
            try:
                self.end_date = datetime.strptime(self.laatste_bericht, '%Y-%m-%d %H:%M:%S')
            except:
                Error = FatalError()
                Error.specific_error(4)
        else:
            self.end_date = ""

    def compute_repair_time(self):
        if self.datum_aangemaakt != "" and self.laatste_bericht != "" and self.status == "Afgesloten":
            response = self.end_date - self.start_date
            response_days = ((response.days * 24) + (response.seconds / 3600))
            self.respons_time = response_days
        else:
            self.respons_time = ""


class Data:
    def __init__(self, records):
        self.records = records
        self.list_persons()
        self.list_periods()
        self.list_causes()
        self.list_root_causes()

    def add_record(self, record):
        self.records.append(record)

    def remove_records(self, record_id):
        del self.records[record_id]

    def print_records(self):
        i = 1
        for record in self.records:
            print("record {} out of {}".format(i, len(self.records)))
            record.print_record()
            i += 1

    def list_persons(self):
        self.personList = []
        for record in self.records:
            if record.klant not in self.personList:
                self.personList.append(record.klant)

    def list_periods(self):
        self.periodList = []
        for record in self.records:
            period = str(record.month) + " " + str(record.year)
            if period not in self.periodList:
                self.periodList.append(period)

    def list_causes(self):
        self.originList = []
        for record in self.records:
            origin = record.root_cause02
            if origin not in self.originList:
                self.originList.append(origin)

    def list_root_causes(self):
        self.root_cause_list = []
        for record in self.records:
            root_cause = record.root_cause
            if root_cause not in self.root_cause_list:
                self.root_cause_list.append(root_cause)

    def filter_person(self, person_name):
        selection = []
        for record in self.records:
            if record.klant == person_name:
                selection.append(record)
        return selection

    def filter_root_cause(self, root_cause_name):
        selection = []
        for record in self.records:
            if record.root_cause == root_cause_name:
                selection.append(record)
        return selection


class Importer:
    def __init__(self, file_name):
        print(">>> starting import")
        self.Error = FatalError()
        self.file_name = file_name
        self.check_file()
        self.read_file()
        self.check_data()
        self.import_data()
        print(">>> import complete")

    def check_file(self):
        print("     # ... validating file")

        if self.file_name[-4:] == ".xls":
            try:
                print("     # ... converting file to csv")
                data = pandas.read_excel(self.file_name)
                self.file_name = self.file_name[:-4] + ".csv"
                data.to_csv(self.file_name, sep=";", index=False)
            except:
                self.Error.specific_error(1)

        elif self.file_name[-4:] == ".csv":
            print("     # ... valid .csv file")

        else:
            print("     # ... invalid file type, file should be of type .csv or .xls")
            self.Error.specific_error(5)

    def read_file(self):
        print("     # ... importing csv")
        try:
            self.data = pandas.read_csv(self.file_name, sep=";", keep_default_na=False)
            self.data = self.data.values.tolist()
        except:
            self.Error.specific_error(2)

    def check_data(self):
        condition_1 = len(self.data[0]) == 21

        if condition_1 is False:
            print("invalid data format, number of columns does not match")
            self.Error.specific_error(3)

    def import_data(self):
        print("     # ... reading data")
        try:
            records = []
            for row in self.data:
                new_record = Record(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9],
                                    row[10],
                                    row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19],
                                    row[20]
                                    )
                records.append(new_record)

            self.Data = Data(records)
        except:
            self.Error.specific_error(3)


class Analyser:

    def __init__(self, file_location, start, stop, ref_start, ref_stop, FM):
        self.start = start
        self.stop = stop
        self.ref_start = ref_start
        self.ref_stop = ref_stop
        self.sep = ";"
        self.file_location = file_location
        self.Error = FatalError()
        self.FM = FM
        self.figures = {}

    def derive_data(self):
        # import data
        try:
            self.Importer = Importer(self.file_location)
        except:
            self.Error.class_error("Importer")

        # filter data on date
        try:
            self.SelectedData = self.select_by_date(self.start, self.stop)

            self.SelectedData.list_periods()
        except:
            self.Error.class_error("Analyser filtering data")

        if len(self.SelectedData.records) == 0:
            self.Error.class_error("zero elements in selection")

        # reference data
        try:
            if self.ref_start != "none" and self.ref_stop != "none":
                self.referenceData = self.select_by_date(self.ref_start, self.ref_stop)
                if len(self.SelectedData.records) == 0:
                    self.Error.class_error("zero elements in refernce selection")

            else:
                self.referenceData = "none"
        except:
            self.Error.class_error("Ananyler method derive_data")



    def make_plots(self):

        # create plotting objects
        print(">>> plotting started")
        print("     # ... ")

        try:
            self.RequestByPeriod = RequestByPeriod(self.SelectedData)
            self.RequestByPeriod.plot()
            self.figures["RequestByPeriod"] = self.RequestByPeriod.fig
        except:
            self.Error.class_error("RequestByPeriod")

        try:
            self.RequestByPerson = RequestByPerson(self.SelectedData)
            self.RequestByPerson.plot()
            self.figures["RequestByPerson"] = self.RequestByPerson.fig
        except:
            self.Error.class_error("RequestByPerson")
        try:
            self.RequestCause = RequestCause(self.SelectedData)
            self.RequestCause.plot()
            self.figures["Requestcause"] = self.RequestCause.fig
        except:
            self.Error.class_error("RequestCause")
        try:
            self.RequestRepairTime = RequestRepairTime(self.SelectedData)
            self.RequestRepairTime.plot()
            self.figures["RequestRepairTime"] = self.RequestRepairTime.fig
        except:
            self.Error.class_error("RequestRepairTime")
        try:
            self.RequestRootCause = RequestRootCause(self.SelectedData)
            self.RequestRootCause.plot()
            self.figures["RequestRootCause"] = self.RequestRootCause.fig
        except:
            self.Error.class_error("RequestRootCause")


        self.RequestByTime = RequestByTime(self.SelectedData, self.referenceData, self.start, self.stop, self.ref_start, self.ref_stop)
        self.RequestByTime.plot()
        self.figures["Compare"] = self.RequestByTime.fig


        print(">>> finished plotting")

    def select_by_date(self, start, stop):
        to_remove = []
        selection = copy.deepcopy(self.Importer.Data)
        for i in range(len(selection.records)):

            if isinstance(self.Importer.Data.records[i].start_date, datetime) \
                          and isinstance(self.Importer.Data.records[i].end_date, datetime):

                if (start <= selection.records[i].start_date < (stop) + dt.timedelta(days=1)) is False:
                    to_remove.append(i)

        to_remove.sort(reverse=True)
        for i in to_remove:
            selection.remove_records(i)
        return selection

    def save_figs(self):
        print(">>> saving figures")
        print("     # ... ")
        try:
            for figure in self.figures:
                self.figures[figure].savefig(self.FM.new_dir + "/" + figure + ".png")
        except:
            self.Error.class_error("Analyser, method: save_figs")
        print(">>> figures saved")


class RequestByPeriod:

    def __init__(self, data):
        self.data = data
        self.count_by_period()
        self.Error = FatalError()

    def count_by_period(self):
        self.person_list = []

        for person in self.data.personList:
            selection = self.data.filter_person(person)
            data = [person, self.data.periodList, np.zeros(len(self.data.periodList)).tolist()]

            for record in selection:
                for i in range(len(data[1])):
                    if record.period == data[1][i]:
                        data[2][i] += 1
            self.person_list.append(data)

    def plot(self):
        try:
            self.fig = plt.figure(num="RequestByPeriod", figsize=(20, 10))
            ax = self.fig.add_subplot(111)

            bottom = np.zeros(len(self.person_list[0][1])).tolist()
            p_list = []
            legend_list = []
            for i in range(len(self.person_list)):
                p = ax.bar(self.person_list[i][1], self.person_list[i][2], bottom=bottom)
                p_list.append(p)
                bottom = list(map(operator.add, bottom, self.person_list[i][2]))
                legend_list.append(self.person_list[i][0])

                for period in range(len(self.person_list[i][1])):
                    if self.person_list[i][2][period] != 0:
                        plt.text(x=period, y=(bottom[period] - (self.person_list[i][2][period] / 2)),
                                 s=str(int(self.person_list[i][2][period])))
            plt.legend(p_list, legend_list)

            if len(self.person_list[0][1]) > 10:
                for tick in ax.get_xticklabels():
                    tick.set_rotation(45)

            ax.set_ylabel("frequetie")
            plt.grid(axis='y')
        except:
            self.Error.class_error("plotting Request by period")


class RequestByPerson:
    def __init__(self, data):
        self.data = data
        self.count_by_person()
        self.Error = FatalError()

    def count_by_person(self):
        self.person_list = []
        for person in self.data.personList:
            count = 0
            for record in self.data.records:
                if record.klant == person:
                    count += 1
            self.person_list.append([person, count])
        self.person_list = np.array(self.person_list)

    def plot(self):
        try:
            self.fig = plt.figure(num="RequestByPerson", figsize=(20, 10))
            ax = self.fig.add_subplot(111)
            size = 0.3
            wedges, texts, percentages = ax.pie(self.person_list[:, 1], autopct='%1.1f%%', radius=1,
                                                wedgeprops=dict(width=size, edgecolor='w'))
            ax.legend(wedges, self.person_list[:, 0], loc=8)
        except:
            self.Error.class_error("plotting RequestByPerson")


class RequestCause:
    def __init__(self, data):
        self.data = data
        self.count_by_origin()
        self.Error = FatalError()

    def count_by_origin(self):
        origin_list = []
        origin_count = []
        for origin in self.data.originList:
            count = 0
            for record in self.data.records:
                if record.root_cause02 == origin:
                    count += 1
            if origin == "":
                origin = "Not specified"
            origin_list.append(origin)
            origin_count.append(count)
        # sort data
        origin_list = np.array(origin_list)
        origin_count = np.array(origin_count)
        self.origin_list = origin_list[origin_count.argsort()]
        self.origin_count = origin_count[origin_count.argsort()]

    def plot(self):
        try:
            self.fig = plt.figure("RequestCause", figsize=(20, 10))
            ax = self.fig.add_subplot(111)
            ax.barh(self.origin_list, self.origin_count)
            for i in range(len(self.origin_list)):
                plt.text(x=self.origin_count[i], y=i, s=str(self.origin_count[i]))
            plt.grid(axis='x')
            plt.xlabel("Aantal meldingen")
            plt.subplots_adjust(left=0.25)
            plt.title("Oorzaak")
        except:
            self.Error.class_error("plotting RequestCause")


class RequestRootCause:

    def __init__(self, data):
        self.data = data
        self.count_by_root_cause()
        self.Error = FatalError()

    def count_by_root_cause(self):
        self.cause_list = []

        for root_cause in self.data.root_cause_list:
            selection = self.data.filter_root_cause(root_cause)
            data = [root_cause, self.data.periodList, np.zeros(len(self.data.periodList)).tolist()]

            for record in selection:
                for i in range(len(data[1])):
                    if record.period == data[1][i]:
                        data[2][i] += 1

            self.cause_list.append(data)

    def plot(self):
        try:
            self.fig = plt.figure("RequestRootCause", figsize=(20, 10))
            ax = self.fig.add_subplot(111)

            bottom = np.zeros(len(self.cause_list[0][1])).tolist()
            p_list = []
            legend_list = []
            for i in range(len(self.cause_list)):
                p = ax.bar(self.cause_list[i][1], self.cause_list[i][2], bottom=bottom)
                p_list.append(p)
                bottom = list(map(operator.add, bottom, self.cause_list[i][2]))
                legend_list.append(self.cause_list[i][0])

                for period in range(len(self.cause_list[0][1])):
                    if self.cause_list[i][2][period] != 0:
                        plt.text(x=period, y=(bottom[period] - (self.cause_list[i][2][period] / 2)),
                                 s=str(int(self.cause_list[i][2][period])))

            plt.legend(p_list, legend_list)
        except:
            self.Error.class_error("plotting RequestRootCause")

        if len(self.cause_list[0][1]) > 10:
            for tick in ax.get_xticklabels():
                tick.set_rotation(45)

        ax.set_ylabel("frequetie")

        plt.grid(axis='y')


class RequestRepairTime:

    # ATTENTION:
    # this graph is produced on a discontinuous semi exponential scale
    # it creates a negative idea/image that there is peak around 2-3-4 months

    def __init__(self, data):
        self.data = data
        self.categorize()
        self.Error = FatalError()

    def categorize(self):
        self.categories = ["1 dag", "2 dagen", "4 dagen,", "1 week", "2 weken", "1 maand", "2 maanden", "4 maanden",
                           "6 maanden", "1 jaar", "> 1 jaar"]
        self.count = np.zeros(len(self.categories))

        for record in self.data.records:
            if record.respons_time != "" and record.status == "Afgesloten":
                if record.respons_time <= 24:
                    self.count[0] += 1
                elif 24 < record.respons_time <= 48:
                    self.count[1] += 1
                elif 48 < record.respons_time <= 96:
                    self.count[2] += 1
                elif 96 < record.respons_time <= 168:
                    self.count[3] += 1
                elif 168 < record.respons_time <= 336:
                    self.count[4] += 1
                elif 336 < record.respons_time <= 744:
                    self.count[5] += 1
                elif 744 < record.respons_time <= 1464:
                    self.count[6] += 1
                elif 1464 < record.respons_time <= 2928:
                    self.count[7] += 1
                elif 2928 < record.respons_time <= 4392:
                    self.count[8] += 1
                elif 4392 < record.respons_time <= 8784:
                    self.count[9] += 1
                elif 8784 < record.respons_time:
                    self.count[10] += 1

    def plot(self):
        try:
            self.fig = plt.figure(num="RequestRepairTime", figsize=(20, 10))
            ax = self.fig.add_subplot(111)
            plt.bar(self.categories, self.count)
            plt.title("Oplossingstijd (Aanmaak ticket - laatste bericht)")
            plt.yticks(np.arange(0, max(self.count) + 5, 5))
            plt.ylabel("Aantal")
            for i in range(len(self.categories)):
                plt.text(x=i, y=self.count[i], s=int(self.count[i]))
            plt.grid(axis='y')

            if len(self.categories) > 10:
                for tick in ax.get_xticklabels():
                    tick.set_rotation(45)
        except:
            self.Error.class_error("plotting ReqeustRepairTime")


class RequestByTime:

    def __init__(self, data, reference_data, start, stop, ref_start, ref_stop):
        self.data = data
        self.reference_data = reference_data
        self.start = start
        self.stop = stop
        self.ref_start = ref_start
        self.ref_stop = ref_stop
        self.prepare_data()
        self.Error = FatalError()

    def prepare_data(self):
        self.values = []
        for record in self.data.records:
            if record.respons_time != "":
                self.values.append(record.respons_time/24)

        if self.ref_start == "none" and self.ref_stop == "none":
            self.ref_values = "none"

        else:
            self.ref_values = []
            for record in self.reference_data.records:
                if record.respons_time != "":
                    self.ref_values.append(record.respons_time/24)

    def plot(self):

        try:
            label_sel = str(self.start.day) + "/" + str(self.start.month) + "/" + str(self.start.year) + " - "\
                        + str(self.stop.day) + "/" + str(self.stop.month) + "/" + str(self.stop.year)

            self.fig = plt.figure(num="Compare", figsize=(20, 10))

            (hist_plot, box_plot) = self.fig.subplots(2, gridspec_kw={"height_ratios": (0.70, 0.30)})


            if self.ref_values != "none":
                label_ref = str(self.ref_start.day) + "/" + str(self.ref_start.month) + "/" + str(
                    self.ref_start.year) + " - " \
                            + str(self.ref_stop.day) + "/" + str(self.ref_stop.month) + "/" + str(self.ref_stop.year)

                bins = np.arange(0, np.max([np.max(self.values), np.max(self.ref_values)]) + 7, 7)
                labels = np.arange(1, np.max([self.ref_values, self.ref_values]) + 1, 1, dtype=int)

                hist_plot.hist([self.ref_values, self.values], bins=bins, label=[label_ref, label_sel],
                               color=['r', 'g'], align='mid')

                box_plot.boxplot([self.values, self.ref_values], vert=False, patch_artist=True,
                                 labels=[label_ref, label_sel])

            if self.ref_values == "none":
                bins = np.arange(0, np.max(np.max(self.values)) + 7, 7)
                labels = np.arange(1, np.max(self.values) + 1, 1, dtype=int)

                hist_plot.hist(self.values, bins=bins, label=label_sel, color=['g'], align='mid')

                box_plot.boxplot(self.values, vert=False, patch_artist=True, labels=[label_sel])

            xticks = np.arange(3.5, len(labels) + 7, 7)
            hist_plot.set_xticks(xticks)
            hist_plot.set_xticklabels(labels)
            hist_plot.set_xlabel("Hersteltijd in weken (Aanmaak ticket - laatste bericht)")
            hist_plot.grid()
            hist_plot.legend()

            box_plot.set_xticks(xticks)
            box_plot.set_xticklabels(labels)
            box_plot.set_xlabel("Hersteltijd in weken (Aanmaak ticket - laatste bericht)")
            box_plot.grid(axis='x')
        except:
            self.Error.class_error("plotting RequestByTime")


class FatalError:
    def __init__(self):
        self.specific_messages = {0: "      # ... something went wrong",
                                  1: "      # ... could not read xls file",
                                  2: "      # ... could not read csv file",
                                  3: "      # ... file does not match expected data",
                                  4: "      # ... unexpected date-time notation",
                                  5: "      # ... no such file",
                                  6: "      # ... failed to filter on date"}

    def specific_error(self, id):
        print("__ERROR__")
        print("      # ... error type: specific error")
        print(self.specific_messages[id])
        self.exit()

    def class_error(self, class_name):
        print("__ERROR__: ")
        print("      # ... error type: class error")
        print("      # ... error in class {}".format(class_name))
        self.exit()

    def exit(self):
        while True:
            print(">>> the error is fatal for the program, close [c]?")
            answer = input()
            if answer == "c":
                print(" bye bye")
                print(5)
                time.sleep(1)
                print(4)
                time.sleep(1)
                print(3)
                time.sleep(1)
                print(2)
                time.sleep(1)
                print(1)
                time.sleep(1)
                sys.exit()


class UserInterface:
    def __init__(self):
        print("   ---------------------------")
        print("   Teamleader export analyser")
        print("   version 1.1.1")
        print("   @author: Willem Boone")
        print("   ---------------------------")
        print("")

    def ask_date_selection(self):
        while True:
            print(">>> do you want to select a specific date range? [y/n]")
            answer = input()
            if answer == "y":
                self.start_date = self.ask_start_date()
                self.end_date = self.ask_end_date()
                if self.end_date < self.start_date:
                    print("     # ... end date should be later then start date")
                else:
                    self.ask_comp_period()
                    break
            elif answer == "n":
                self.end_date = datetime.strptime('2100-12-31', '%Y-%m-%d')
                self.start_date = datetime.strptime('1900-12-31', '%Y-%m-%d')
                self.ref_start = "none"
                self.ref_stop = "none"
                break
            else:
                print("     # ... invalid response")

    def ask_comp_period(self):
        while True:
            print(">>> do you want to select a period to compare with? [y/n]")
            answer = input()
            if answer == "y":
                while True:
                    self.ref_start = self.ask_start_date()
                    if self.ref_start >= self.start_date:
                        print("     # ... the start of the comparison period should not be within the period of interest")
                    else:
                        break

                while True:
                    self.ref_stop = self.ask_end_date()
                    print(self.ref_stop >= self.start_date)
                    if self.ref_stop >= self.start_date:
                        print("     # ... the end of the comparison period should not be within the period of interest")
                    else:
                        break

                if self.ref_stop < self.ref_start:
                    print("     # ... end date should be later then start date")
                else:
                    break
            elif answer == "n":
                self.ref_start = "none"
                self.ref_stop = "none"
                break
            else:
                print("     # ... invalid response")

    def ask_start_date(self):
        while True:
            print("     # ... start date? [yyyy-mm-dd]")
            start = input()
            try:
                start_date = datetime.strptime(start, '%Y-%m-%d')
            except:
                print("     # ... invalid date")
            else:
                return start_date

    def ask_end_date(self):
        while True:
            print("     # ... End date? [yyyy-mm-dd]")
            end = input()
            try:
                end_date = datetime.strptime(end, '%Y-%m-%d')
            except:
                print("     # ... invalid date")
            else:
                return end_date

    def ask_file_location(self):
        while True:
            print(">>> what is the location of the export file?")
            answer = input()
            if os.path.isfile(answer):
                self.file_location = answer
                break
            else:
                print("     # this is not a valid file location")

    def finish(self):
        print(">>> program has finished")
        while True:
            print("     # ... terminate program? [y/n]")
            answer = input()
            if answer == "y":
                print(">>> bye bye")
                print(5)
                time.sleep(1)
                print(4)
                time.sleep(1)
                print(3)
                time.sleep(1)
                print(2)
                time.sleep(1)
                print(1)
                time.sleep(1)
                break
            elif answer == "n":
                print("     # i finished my job, please let me go to sleep now")
            else:
                print("     # cannot understand command")


class Statistics:
    def __init__(self, selected_data, ref_data, start, stop, ref_start, ref_stop, FM):
        self.FM = FM
        self.selected_data = selected_data
        self.ref_data = ref_data
        self.start = start
        self.stop = stop
        self.ref_start = ref_start
        self.ref_stop = ref_stop
        self.dict_sel = {}
        self.dict_ref = {}
        self.dict_comp = {}
        self.Error = FatalError()

    def calculate_statistics(self):
        print(">>> calculating statistics")
        print("      # ...")
        self.selection_statistics()

        if self.ref_data != "none":
            self.reference_statistics()
            self.write_array()
            self.write_csv()

        self.write_dict()

    def selection_statistics(self):
        try:
            # number of tickets
            self.ticket_numbers = len(self.selected_data.records)

            # status of tickets
            self.afgesloten = 0
            self.third_line = 0
            self.second_line = 0
            self.first_line = 0
            self.waiting_client = 0
            self.undefined_status = 0

            for record in self.selected_data.records:
                if record.status == "Afgesloten":
                    self.afgesloten += 1
                elif record.status == "3d Line":
                    self.third_line += 1
                elif record.status == "2nd Line":
                    self.second_line += 1
                elif record.status == "1st Line":
                    self.first_line += 1
                elif record.status == "Wachten op klant":
                    self.waiting_client += 1
                else:
                    self.undefined_status += 1

            # average repair time
            sum = 0
            count = 0
            for record in self.selected_data.records:
                if record.respons_time != "":
                    sum += record.respons_time
                    count += 1
            self.average_repair_time = np.round((sum / count) / 24)

            # cause statistics
            self.cause_bug = 0
            self.cause_config = 0
            self.cause_knowl = 0
            self.cause_undet = 0
            self.cause_ext = 0
            self.cause_int = 0
            self.cause_missing = 0
            self.cause_inf = 0
            self.cause_notspec = 0

            for record in self.selected_data.records:
                if record.root_cause02 == "Bug software":
                    self.cause_bug += 1
                elif record.root_cause02 == "Configuration request":
                    self.cause_config += 1
                elif record.root_cause02 == "Lack of functional knowledge":
                    self.cause_knowl += 1
                elif record.root_cause02 == "Undetermined":
                    self.cause_undet += 1
                elif record.root_cause02 == "External service outage":
                    self.cause_ext += 1
                elif record.root_cause02 == "Internal service outage":
                    self.cause_int += 1
                elif record.root_cause02 == "Missing Feature":
                    self.cause_missing += 1
                elif record.root_cause02 == "Wrong information from customer":
                    self.cause_inf += 1
                elif record.root_cause02 == "Not specified":
                    self.cause_notspec += 1

            # root cause
            self.rc_geo = 0
            self.rc_notsp = 0
            self.rc_cust = 0
            self.rc_undet = 0
            self.rc_ext = 0

            for record in self.selected_data.records:
                if record.root_cause == "Geosparc":
                    self.rc_geo += 1
                elif record.root_cause == "Not specified":
                    self.rc_notsp += 1
                elif record.root_cause == "Customer":
                    self.rc_cust += 1
                elif record.root_cause == "Undetermined":
                    self.rc_undet += 1
                elif record.root_cause == "External":
                    self.rc_ext += 1

        except:
            self.Error.class_error("Statics: selection statistics")

    def reference_statistics(self):
        try:
            # number of tickets
            self.ticket_numbers_ref = len(self.ref_data.records)

            # status of tickets
            self.afgesloten_ref = 0
            self.third_line_ref = 0
            self.second_line_ref = 0
            self.first_line_ref = 0
            self.waiting_client_ref = 0
            self.undefined_status_ref = 0

            for record in self.ref_data.records:
                if record.status == "Afgesloten":
                    self.afgesloten_ref += 1
                elif record.status == "3d Line":
                    self.third_line_ref += 1
                elif record.status == "2nd Line":
                    self.second_line_ref += 1
                elif record.status == "1st Line":
                    self.first_line_ref += 1
                elif record.status == "Wachten op klant":
                    self.waiting_client_ref += 1
                else:
                    self.undefined_status_ref += 1

            # average repair time
            sum = 0
            count = 0
            for record in self.ref_data.records:
                if record.respons_time != "":
                    sum += record.respons_time
                    count += 1
            self.average_repair_time_ref = np.round((sum / count) / 24)

            # cause statistics
            self.cause_bug_ref = 0
            self.cause_config_ref = 0
            self.cause_knowl_ref = 0
            self.cause_undet_ref = 0
            self.cause_ext_ref = 0
            self.cause_int_ref = 0
            self.cause_missing_ref = 0
            self.cause_inf_ref = 0
            self.cause_notspec_ref = 0

            for record in self.ref_data.records:
                if record.root_cause02 == "Bug software":
                    self.cause_bug_ref += 1
                elif record.root_cause02 == "Configuration request":
                    self.cause_config_ref += 1
                elif record.root_cause02 == "Lack of functional knowledge":
                    self.cause_knowl_ref += 1
                elif record.root_cause02 == "Undetermined":
                    self.cause_undet_ref += 1
                elif record.root_cause02 == "External service outage":
                    self.cause_ext_ref += 1
                elif record.root_cause02 == "Internal service outage":
                    self.cause_int_ref += 1
                elif record.root_cause02 == "Missing Feature":
                    self.cause_missing_ref += 1
                elif record.root_cause02 == "Wrong information from customer":
                    self.cause_inf_ref += 1
                elif record.root_cause02 == "":
                    self.cause_notspec_ref += 1

            # root cause
            self.rc_geo_ref = 0
            self.rc_notsp_ref = 0
            self.rc_cust_ref = 0
            self.rc_undet_ref = 0
            self.rc_ext_ref = 0

            for record in self.ref_data.records:
                if record.root_cause == "Geosparc":
                    self.rc_geo_ref += 1
                elif record.root_cause == "Not specified":
                    self.rc_notsp_ref += 1
                elif record.root_cause == "Customer":
                    self.rc_cust_ref += 1
                elif record.root_cause == "Undetermined":
                    self.rc_undet_ref += 1
                elif record.root_cause == "External":
                    self.rc_ext_ref += 1

        except:
            self.Error.class_error("Statics: selection statistics")

    def compare(self, x, y):
        if x == 0 and y != 0:
            z = -100
        elif y == 0 and x != 0:
            z =  100
        elif y == 0 and x == 0:
            z = 0
        else:
            z = np.round((x-y)/y*100)
        return z

    def write_array(self):
        self.array = [
            ["status: Afgesloten", self.afgesloten, self.afgesloten/self.ticket_numbers, self.afgesloten_ref, self.afgesloten_ref/self.ticket_numbers_ref],
            ["status: 3d Line", self.third_line, self.third_line/self.ticket_numbers, self.third_line_ref, self.third_line_ref/self.ticket_numbers_ref],
            ["status: 2nd Line", self.second_line, self.second_line/self.ticket_numbers, self.second_line_ref, self.second_line_ref/self.ticket_numbers_ref],
            ["status: 1st Line", self.first_line, self.first_line/self.ticket_numbers, self.first_line_ref, self.first_line_ref/self.ticket_numbers_ref],
            ["status: wachten op klant", self.waiting_client, self.waiting_client/self.ticket_numbers, self.waiting_client_ref, self.waiting_client_ref/self.ticket_numbers_ref],
            ["status: ongekend", self.undefined_status, self.undefined_status/self.ticket_numbers, self.undefined_status_ref, self.undefined_status_ref/self.ticket_numbers_ref],
            ["oorzaak: Bug software", self.cause_bug, self.cause_bug/self.ticket_numbers, self.cause_bug_ref, self.cause_bug_ref/self.ticket_numbers_ref],
            ["oorzaak: Configuration request", self.cause_config, self.cause_config/self.ticket_numbers, self.cause_config_ref, self.cause_config_ref/self.ticket_numbers_ref],
            ["oorzaak: Lack of functional knowledge", self.cause_knowl, self.cause_knowl/self.ticket_numbers, self.cause_knowl_ref, self.cause_knowl_ref/self.ticket_numbers_ref],
            ["oorzaak: Undetermined", self.cause_undet, self.cause_undet/self.ticket_numbers, self.cause_undet_ref, self.cause_undet_ref/self.ticket_numbers_ref],
            ["oorzaak: External service outage", self.cause_ext, self.cause_ext/self.ticket_numbers, self.cause_ext_ref, self.cause_ext_ref/self.ticket_numbers_ref],
            ["oorzaak: Internal service outage", self.cause_int, self.cause_int/self.ticket_numbers, self.cause_int_ref, self.cause_int_ref/self.ticket_numbers_ref],
            ["oorzaak: Missing Feature", self.cause_missing, self.cause_missing/self.ticket_numbers, self.cause_missing_ref, self.cause_missing_ref/self.ticket_numbers_ref],
            ["oorzaak: Wrong information from customer", self.cause_inf, self.cause_inf/self.ticket_numbers, self.cause_inf_ref, self.cause_inf_ref/self.ticket_numbers_ref],
            ["oorzaak: Not specified", self.cause_notspec, self.cause_notspec/self.ticket_numbers, self.cause_notspec_ref, self.cause_notspec_ref/self.ticket_numbers_ref],
            ["root cause: Geosparc", self.rc_geo, self.rc_geo/self.ticket_numbers, self.rc_geo_ref, self.rc_geo_ref/self.ticket_numbers_ref],
            ["root cause: Customer", self.rc_cust, self.rc_cust/self.ticket_numbers, self.rc_cust_ref, self.rc_cust_ref/self.ticket_numbers_ref],
            ["root cause: Exteral", self.rc_ext, self.rc_ext/self.ticket_numbers, self.rc_ext_ref, self.rc_ext_ref/self.ticket_numbers_ref],
            ["root cause: Undetermined", self.rc_undet, self.rc_undet/self.ticket_numbers, self.rc_undet_ref, self.rc_undet_ref/self.ticket_numbers_ref],
            ["root cause: not specified", self.rc_notsp, self.rc_notsp/self.ticket_numbers, self.rc_notsp_ref, self.rc_notsp_ref/self.ticket_numbers_ref]
        ]
        for row in self.array:
            row.append(self.compare(row[1], row[3]))

    def write_dict(self):

        # selection dict
        self.dict_sel["aantal tickets"] = self.ticket_numbers
        self.dict_sel["gemiddelde hersteltijd (dagen)"] = self.average_repair_time
        self.dict_sel["status: Afgesloten"] = self.afgesloten
        self.dict_sel["status: 3d Line"] = self.third_line
        self.dict_sel["status: 2nd Line"] = self.second_line
        self.dict_sel["status: 1st Line"] = self.first_line
        self.dict_sel["status: wachten op klant"] = self.waiting_client
        self.dict_sel["status: ongekend"] = self.undefined_status
        self.dict_sel["oorzaak: Bug software"] = self.cause_bug
        self.dict_sel["oorzaak: Configuration request"] = self.cause_config
        self.dict_sel["oorzaak: Lack of functional knowledge"] = self.cause_knowl
        self.dict_sel["oorzaak: Undetermined"] = self.cause_undet
        self.dict_sel["oorzaak: External service outage"] = self.cause_ext
        self.dict_sel["oorzaak: Internal service outage"] = self.cause_int
        self.dict_sel["oorzaak: Missing Feature"] = self.cause_missing
        self.dict_sel["oorzaak: Wrong information from customer"] = self.cause_inf
        self.dict_sel["oorzaak: Not specified"] = self.cause_notspec
        self.dict_sel["root cause: Geosparc"] = self.rc_geo
        self.dict_sel["root cause: Customer"] = self.rc_cust
        self.dict_sel["root cause: Exteral"] = self.rc_ext
        self.dict_sel["root cause: Undetermined"] = self.rc_undet
        self.dict_sel["root cause: not specified"] = self.rc_notsp

        # reference dict
        if self.ref_data != "none":
            self.dict_ref["aantal tickets"] = self.ticket_numbers_ref
            self.dict_ref["gemiddelde hersteltijd (dagen)"] = self.average_repair_time_ref
            self.dict_ref["status: Afgesloten"] = self.afgesloten_ref
            self.dict_ref["status: 3d Line"] = self.third_line_ref
            self.dict_ref["status: 2nd Line"] = self.second_line_ref
            self.dict_ref["status: 1st Line"] = self.first_line_ref
            self.dict_ref["status: wachten op klant"] = self.waiting_client_ref
            self.dict_ref["status: ongekend"] = self.undefined_status_ref
            self.dict_ref["oorzaak: Bug software"] = self.cause_bug_ref
            self.dict_ref["oorzaak: Configuration request"] = self.cause_config_ref
            self.dict_ref["oorzaak: Lack of functional knowledge"] = self.cause_knowl_ref
            self.dict_ref["oorzaak: Undetermined"] = self.cause_undet_ref
            self.dict_ref["oorzaak: External service outage"] = self.cause_ext_ref
            self.dict_ref["oorzaak: Internal service outage"] = self.cause_int_ref
            self.dict_ref["oorzaak: Missing Feature"] = self.cause_missing_ref
            self.dict_ref["oorzaak: Wrong information from customer"] = self.cause_inf_ref
            self.dict_ref["oorzaak: Not specified"] = self.cause_notspec_ref
            self.dict_ref["root cause: Geosparc"] = self.rc_geo_ref
            self.dict_ref["root cause: Customer"] = self.rc_cust_ref
            self.dict_ref["root cause: Exteral"] = self.rc_ext_ref
            self.dict_ref["root cause: Undetermined"] = self.rc_undet_ref
            self.dict_ref["root cause: not specified"] = self.rc_notsp_ref

            # relative

    def write_txt(self):
        try:
            print(">>> writing summary text file")
            print("      # ...")

            txt = open(self.FM.new_dir + "/statistics.txt", "w+")

            # general message
            txt.write("Analyser txt file output\n")
            txt.write("------------------------\n")
            txt.write("geanalyseerd bestand: {}\n".format(self.FM.base_name))
            txt.write("selectie periode: {} - {}\n".format(self.start, self.stop))
            if self.ref_data != "none":
                txt.write("referentie periode: {} - {}\n".format(self.ref_start, self.ref_stop))
            txt.write("------------------------\n")

            # selection data
            txt.write("--Selectie periode--\n")
            for key in self.dict_sel:
                txt.write("{} = {} \n".format(key, self.dict_sel[key]))
            txt.write("\n")

            if self.selected_data != "none":
                # reference data
                txt.write("--Referentie periode--\n")
                for key in self.dict_ref:
                    txt.write("{} = {} \n".format(key, self.dict_ref[key]))
                txt.write("\n")

            txt.close()

        except:
            self.Error.class_error("Write txt file")

    def write_csv(self):

        print(">>> writing csv comparison file")
        print("      # ...")

        with open(self.FM.new_dir + "/comparison.csv", mode='w', newline ='') as csv_file:
            csv_writer = csv.writer(csv_file, delimiter=';')
            csv_writer.writerow(["variabele", "referentie", "%", "selectie", "%", "toename (%)"])
            csv_writer.writerow(["aantal tickets", self.ticket_numbers_ref, 100, self.ticket_numbers, 100, self.compare(self.ticket_numbers, self.ticket_numbers_ref)])
            csv_writer.writerow(["gemiddelde hersteltijd (dagen)", self.average_repair_time_ref, "", self.average_repair_time, "", self.compare(self.average_repair_time, self.average_repair_time_ref)])
            for row in self.array:
                csv_writer.writerow([row[0], row[3], np.round(row[4] * 100), row[1], np.round(row[2] * 100), row[5]])


class FileManagement:

    def __init__(self, file_name):
        self.file_name = file_name
        self.dir = os.path.dirname(self.file_name)
        self.base_name = os.path.basename(self.file_name)
        self.base_name_clean = os.path.splitext(self.base_name)[0]
        self.new_dir = str(self.dir) + "/analyse_" + str(self.base_name_clean)
        try:
            os.mkdir(self.new_dir)
        except:
            pass


class Program:
    def __init__(self):
        self.Error = FatalError()
        print("PROGRAM RUNNING...")

    def test_UI(self):
        try:
            UI = UserInterface()
            UI.ask_file_location()
            UI.ask_date_selection()
            FM = FileManagement(UI.file_location)
            DemoAnalyser = Analyser(UI.file_location, UI.start_date, UI.end_date, UI.ref_start, UI.ref_stop, FM)
            UI.finish()
            print("UI test successful")
        except:
            print("UI test failed")

    def test_analyser(self, file):
        try:
            start = datetime.strptime('2018-09-01', '%Y-%m-%d')
            stop = datetime.strptime('2019-08-31', '%Y-%m-%d')
            ref_start = datetime.strptime('2017-09-01', '%Y-%m-%d')
            ref_stop = datetime.strptime('2018-08-31', '%Y-%m-%d')
            ref_start = "none"
            ref_stop = "none"
            FM = FileManagement(file)
            DemoAnalyser = Analyser(file, start, stop, ref_start, ref_stop, FM)
            DemoAnalyser.derive_data()
            DemoAnalyser.make_plots()
            DemoAnalyser.save_figs()
            print("analyser test successful")
        except:
            print("analyser test failed")

    def execute(self):
        try:
            UI = UserInterface()
            UI.ask_file_location()
            UI.ask_date_selection()
            FM = FileManagement(UI.file_location)
            DemoAnalyser = Analyser(UI.file_location, UI.start_date, UI.end_date, UI.ref_start, UI.ref_stop, FM)
            DemoAnalyser.derive_data()
            DemoAnalyser.make_plots()
            DemoAnalyser.save_figs()
            stat = Statistics(DemoAnalyser.SelectedData, DemoAnalyser.referenceData, UI.start_date, UI.end_date,
                              UI.ref_start, UI.ref_stop, FM)
            stat.calculate_statistics()
            stat.write_txt()
            UI.finish()
        except:
            self.Error.class_error("Undefined error")

    def test_program(self, file):

        start = datetime.strptime('2018-09-01', '%Y-%m-%d')
        stop = datetime.strptime('2019-08-31', '%Y-%m-%d')
        ref_start = datetime.strptime('2017-09-01', '%Y-%m-%d')
        ref_stop = datetime.strptime('2018-08-31', '%Y-%m-%d')
        #ref_start = "none"
        #ref_stop = "none"
        FM = FileManagement(file)
        DemoAnalyser = Analyser(file, start, stop, ref_start, ref_stop, FM)
        DemoAnalyser.derive_data()
        DemoAnalyser.make_plots()
        DemoAnalyser.save_figs()
        stat = Statistics(DemoAnalyser.SelectedData, DemoAnalyser.referenceData, start, stop, ref_start, ref_stop, FM)
        stat.calculate_statistics()
        stat.write_txt()
        stat.write_csv()


prog = Program()
prog.execute()
