#########################################################################################################################
#                                                                                                                       #
#   Beschreibung:                                                                                                       #
#   Das Script liest eine Excel Tabelle ein und errechnet alle benötigten Mindest-, Maximal- und Durchschnittswerte.    #
#                                                                                                                       #                                                       
#   Autor: Patrik Faust                                                                                                 #
#                                                                                                                       #
#########################################################################################################################

import pandas as pd 

class Measurements(object):

    def __init__(self, excel_file):
        self.excel_file = excel_file

    def _get_avg_from_sheet(self, sheet, column):
        numeric_sheet = pd.to_numeric(sheet[column], errors='coerce')
        return numeric_sheet.mean().round(2)

    def _get_min_from_sheet(self, sheet, column):
        numeric_sheet = pd.to_numeric(sheet[column], errors='coerce')
        return numeric_sheet.min()
    
    def _get_max_from_sheet(self, sheet, column):
        numeric_sheet = pd.to_numeric(sheet[column], errors='coerce')
        return numeric_sheet.max()

    def _get_index_2_4(self, sheet):
        index_list = (sheet["Unnamed: 1"].to_list())
        for x in range(len(index_list)):
            if (index_list[x] == "IEEE 802.11a/n/ac/ax (40 MHz)"):
                return x

    def _get_sheet_2_4(self, sheet):
        performance_sheet_2_4 = (sheet.head(self._get_index_2_4(sheet)))
        return performance_sheet_2_4
    
    def _get_sheet_5(self, sheet):
        performance_sheet_5 = (sheet.tail(len(sheet) - self._get_index_2_4(sheet)))
        return performance_sheet_5
    
    def _get_best_case(self, sheet_req_frequencyband):
        list_best_case_worst_case = (sheet_req_frequencyband["Unnamed: 2"].to_list())
        for x in range(len(list_best_case_worst_case)):
            if list_best_case_worst_case[x] == "Worst-Case":
                index_worst_case = x
        return sheet_req_frequencyband.head(index_worst_case)
    
    def _get_worst_case(self, sheet_req_frequencyband):
        list_best_case_worst_case = (sheet_req_frequencyband["Unnamed: 2"].to_list())
        for x in range(len(list_best_case_worst_case)):
            if list_best_case_worst_case[x] == "Worst-Case":
                index_worst_case = x
        worst_case_range = len(list_best_case_worst_case) - index_worst_case
        return sheet_req_frequencyband.tail(worst_case_range)

class Perfomance(Measurements):

    rssi_column = "Unnamed: 5"
    uplink_column = "Unnamed: 6"
    downlink_column = "Unnamed: 7"
    connection_time_column = "Unnamed: 8"

    def __init__(self, excel_file):
        super().__init__(excel_file)
        self.performance_sheet = pd.read_excel(self.excel_file, sheet_name="Performance", index_col=0)
        self.sheet_2_4 = self._get_sheet_2_4(self.performance_sheet)
        self.sheet_5 = self._get_sheet_5(self.performance_sheet)
        self.best_case_2_4 = self._get_best_case(self.sheet_2_4)
        self.best_case_5 = self._get_best_case(self.sheet_5)
        self.worst_case_2_4 = self._get_worst_case(self.sheet_2_4)
        self.worst_case_5 = self._get_worst_case(self.sheet_5)

        # Verbindungsaufbauzeit Min/Max/Avg über beide Frequenzbänder
        self.min_connection_time_2_4_5 = int(self._get_min_from_sheet(self.performance_sheet, self.connection_time_column))
        self.max_connection_time_2_4_5 = int(self._get_max_from_sheet(self.performance_sheet, self.connection_time_column))
        self.avg_connection_time_2_4_5 = self._get_avg_from_sheet(self.performance_sheet, self.connection_time_column)

        # Verbindungsaufbauzeit Min/Max/Avg im best-case in dem 2,4 GHz Frequenzband
        self.min_best_case_connection_time_2_4 = int(self._get_min_from_sheet(self.best_case_2_4, self.connection_time_column))
        self.max_best_case_connection_time_2_4 = int(self._get_max_from_sheet(self.best_case_2_4, self.connection_time_column))
        self.avg_best_case_connection_time_2_4 = self._get_avg_from_sheet(self.best_case_2_4, self.connection_time_column)

        # Verbindungsaufbauzeit Min/Max/Avg im worst-case in dem 2,4 GHz Frequenzband
        self.min_worst_case_connection_time_2_4 = int(self._get_min_from_sheet(self.worst_case_2_4, self.connection_time_column))
        self.max_worst_case_connection_time_2_4 = int(self._get_max_from_sheet(self.worst_case_2_4, self.connection_time_column))
        self.avg_worst_case_connection_time_2_4 = self._get_avg_from_sheet(self.worst_case_2_4, self.connection_time_column)

        # Verbindungsaufbauzeit Min/Max/Avg im best-case in dem 5 GHz Frequenzband
        self.min_best_case_connection_time_5 = int(self._get_min_from_sheet(self.best_case_5, self.connection_time_column))
        self.max_best_case_connection_time_5 = int(self._get_max_from_sheet(self.best_case_5, self.connection_time_column))
        self.avg_best_case_connection_time_5 = self._get_avg_from_sheet(self.best_case_5, self.connection_time_column)

        # Verbindungsaufbauzeit Min/Max/Avg im worst-case in dem 5 GHz Frequenzband
        self.min_worst_case_connection_time_5 = int(self._get_min_from_sheet(self.worst_case_5, self.connection_time_column))
        self.max_worst_case_connection_time_5 = int(self._get_max_from_sheet(self.worst_case_5, self.connection_time_column))
        self.avg_worst_case_connection_time_5 = self._get_avg_from_sheet(self.worst_case_5, self.connection_time_column)

        # Uplink Datendurchsatz Min/Max/Avg im best-case in dem 2,4 GHz Frequenzband
        self.min_best_case_uplink_2_4 = self._get_min_from_sheet(self.best_case_2_4, self.uplink_column)
        self.max_best_case_uplink_2_4 = self._get_max_from_sheet(self.best_case_2_4, self.uplink_column)
        self.avg_best_case_uplink_2_4 = self._get_avg_from_sheet(self.best_case_2_4, self.uplink_column)

        # Uplink Datendurchsatz Min/Max/Avg im worst-case in dem 2,4 GHz Frequenzband
        self.min_worst_case_uplink_2_4 = self._get_min_from_sheet(self.worst_case_2_4, self.uplink_column)
        self.max_worst_case_uplink_2_4 = self._get_max_from_sheet(self.worst_case_2_4, self.uplink_column)
        self.avg_worst_case_uplink_2_4 = self._get_avg_from_sheet(self.worst_case_2_4, self.uplink_column)

        # Downlink Datendurchsatz Min/Max/Avg im best-case in dem 2,4 GHz Frequenzband
        self.min_best_case_downlink_2_4 = self._get_min_from_sheet(self.best_case_2_4, self.downlink_column)
        self.max_best_case_downlink_2_4 = self._get_max_from_sheet(self.best_case_2_4, self.downlink_column)
        self.avg_best_case_downlink_2_4 = self._get_avg_from_sheet(self.best_case_2_4, self.downlink_column)

        # Downlink Datendurchsatz Min/Max/Avg im worst-case in dem 2,4 GHz Frequenzband
        self.min_worst_case_downlink_2_4 = self._get_min_from_sheet(self.worst_case_2_4, self.downlink_column)
        self.max_worst_case_downlink_2_4 = self._get_max_from_sheet(self.worst_case_2_4, self.downlink_column)
        self.avg_worst_case_downlink_2_4 = self._get_avg_from_sheet(self.worst_case_2_4, self.downlink_column)

        # Uplink Datendurchsatz Min/Max/Avg im best-case in dem 5 GHz Frequenzband
        self.min_best_case_uplink_5 = self._get_min_from_sheet(self.best_case_5, self.uplink_column)
        self.max_best_case_uplink_5 = self._get_max_from_sheet(self.best_case_5, self.uplink_column)
        self.avg_best_case_uplink_5 = self._get_avg_from_sheet(self.best_case_5, self.uplink_column)

        # Uplink Datendurchsatz Min/Max/Avg im worst-case in dem 5 GHz Frequenzband
        self.min_worst_case_uplink_5 = self._get_min_from_sheet(self.worst_case_5, self.uplink_column)
        self.max_worst_case_uplink_5 = self._get_max_from_sheet(self.worst_case_5, self.uplink_column)
        self.avg_worst_case_uplink_5 = self._get_avg_from_sheet(self.worst_case_5, self.uplink_column)

        # Downlink Datendurchsatz Min/Max/Avg im best-case in dem 5 GHz Frequenzband
        self.min_best_case_downlink_5 = self._get_min_from_sheet(self.best_case_5, self.downlink_column)
        self.max_best_case_downlink_5 = self._get_max_from_sheet(self.best_case_5, self.downlink_column)
        self.avg_best_case_downlink_5 = self._get_avg_from_sheet(self.best_case_5, self.downlink_column)

        # Downlink Datendurchsatz Min/Max/Avg im worst-case in dem 5 GHz Frequenzband
        self.min_worst_case_downlink_5 = self._get_min_from_sheet(self.worst_case_5, self.downlink_column)
        self.max_worst_case_downlink_5 = self._get_max_from_sheet(self.worst_case_5, self.downlink_column)
        self.avg_worst_case_downlink_5 = self._get_avg_from_sheet(self.worst_case_5, self.downlink_column)

class Roaming(Measurements):

    package_loss_column = "Unnamed: 7"
    roaming_time_column = "Unnamed: 8"

    def __init__(self, excel_file):
        super().__init__(excel_file)
        self.roaming_sheet = self._fix_package_loss(pd.read_excel(self.excel_file, sheet_name="Roaming", index_col=0))
        self.sheet_2_4 = self._get_sheet_2_4(self.roaming_sheet)
        self.sheet_5 = self._get_sheet_5(self.roaming_sheet)

        # Durchschnittliche Roamingzeit und Paketverlust über beide Frequenzbänder
        self.avg_roaming_time_2_4_5 = self._get_avg_from_sheet(self.roaming_sheet, self.roaming_time_column)
        self.avg_packet_loss_2_4_5 = self._get_avg_from_sheet(self.roaming_sheet, self.package_loss_column)

        # Roamingzeit Min/Max/Avg in dem 2,4 GHz Frequenzband
        self.min_roaming_time_2_4 = int(self._get_min_from_sheet(self.sheet_2_4, self.roaming_time_column))
        self.max_roaming_time_2_4 = int(self._get_max_from_sheet(self.sheet_2_4, self.roaming_time_column))
        self.avg_roaming_time_2_4 = self._get_avg_from_sheet(self.sheet_2_4, self.roaming_time_column)

        # Roamingzeit Min/Max/Avg in dem 5 GHz Frequenzband
        self.min_roaming_time_5 = int(self._get_min_from_sheet(self.sheet_5, self.roaming_time_column))
        self.max_roaming_time_5 = int(self._get_max_from_sheet(self.sheet_5, self.roaming_time_column))
        self.avg_roaming_time_5 = self._get_avg_from_sheet(self.sheet_5, self.roaming_time_column)

        # Paketverlust Min/Max/Avg in dem 2,4 GHz Frequenzband
        self.min_packet_loss_2_4 = self._get_min_from_sheet(self.sheet_2_4, self.package_loss_column)
        self.max_packet_loss_2_4 = self._get_max_from_sheet(self.sheet_2_4, self.package_loss_column)
        self.avg_packet_loss_2_4 = self._get_avg_from_sheet(self.sheet_2_4, self.package_loss_column)

        # Paketverlust Min/Max/Avg in dem 5 GHz Frequenzband
        self.min_packet_loss_5 = self._get_min_from_sheet(self.sheet_5, self.package_loss_column)
        self.max_packet_loss_5 = self._get_max_from_sheet(self.sheet_5, self.package_loss_column)
        self.avg_packet_loss_5 = self._get_avg_from_sheet(self.sheet_5, self.package_loss_column)

    def _fix_package_loss(self, sheet):
        # Löscht alle nicht numerischen Werte aus der Spalte
        sheet[self.package_loss_column] = pd.to_numeric(sheet[self.package_loss_column], errors='coerce')
        # Multipliziert alle Werte in der Spalte mit 100
        sheet[self.package_loss_column] = sheet[self.package_loss_column].apply(lambda x: x*100)
        # Rundet auf 2 Stellen nach dem Komma
        sheet[self.package_loss_column] = sheet[self.package_loss_column].apply(lambda y: round(y, 2))
        fixed_sheet = sheet
        return fixed_sheet

