from docxtpl import DocxTemplate
import re
import excel_reader

class WordWriter(object):
    def __init__(self, word_template, perf_obj: excel_reader.Perfomance, roam_obj: excel_reader.Roaming):
        self.doc = DocxTemplate(word_template)
        self.perf_obj = perf_obj
        self.roam_obj = roam_obj

    def _write_document(self):
        # wenn mehrere Infrastrukturen gemessen wurden, müssen hier noch eine if Verzweigung implementiert werden z.B. if ("Cisco 9120 in infrastrukur_list")
        context =   {
                    # Schreibt die Verbindungsaufbauzeiten über beide Frequenzbänder in das Word Template (Kapitel Fazit und Empfehlungen)
                    'min_best_case_connection_time_2_4_5_cisco9120' : self._notation_helper(self.perf_obj.min_connection_time_2_4_5),
                    'max_best_case_connection_time_2_4_5_cisco9120' : self._notation_helper(self.perf_obj.max_connection_time_2_4_5),
                    'avg_best_case_connection_time_2_4_5_cisco9120' : self._notation_helper(self.perf_obj.avg_connection_time_2_4_5),

                    # Schreibt die Verbindungsaufbauzeiten für das 2,4 GHz Frequenzband in das Word Template (Kapitel Performance Analyse)
                    'min_best_case_connection_time_2_4_cisco9120' : self._notation_helper(self.perf_obj.min_best_case_connection_time_2_4),
                    'max_best_case_connection_time_2_4_cisco9120' : self._notation_helper(self.perf_obj.max_best_case_connection_time_2_4),
                    'avg_best_case_connection_time_2_4_cisco9120' : self._notation_helper(self.perf_obj.avg_best_case_connection_time_2_4),
                    'min_worst_case_connection_time_2_4_cisco9120' : self._notation_helper(self.perf_obj.min_worst_case_connection_time_2_4),
                    'max_worst_case_connection_time_2_4_cisco9120' : self._notation_helper(self.perf_obj.max_worst_case_connection_time_2_4),
                    'avg_worst_case_connection_time_2_4_cisco9120' : self._notation_helper(self.perf_obj.avg_worst_case_connection_time_2_4),

                    # Schreibt die Verbindungsaufbauzeiten für das 5 GHz Frequenzband in das Word Template (Kapitel Performance Analyse)
                    'min_best_case_connection_time_5_cisco9120' : self._notation_helper(self.perf_obj.min_best_case_connection_time_5),
                    'max_best_case_connection_time_5_cisco9120' : self._notation_helper(self.perf_obj.max_best_case_connection_time_5),
                    'avg_best_case_connection_time_5_cisco9120' : self._notation_helper(self.perf_obj.avg_best_case_connection_time_5),
                    'min_worst_case_connection_time_5_cisco9120' : self._notation_helper(self.perf_obj.min_worst_case_connection_time_5),
                    'max_worst_case_connection_time_5_cisco9120' : self._notation_helper(self.perf_obj.max_worst_case_connection_time_5),
                    'avg_worst_case_connection_time_5_cisco9120' : self._notation_helper(self.perf_obj.avg_worst_case_connection_time_5),

                    # Schreibt die Datendurchsatzwerte im Uplink in dem 2,4 GHz Frequenzband in das Word Template (Kapitel Performance Analyse)
                    'min_best_case_uplink_2_4_cisco9120' : self._notation_helper(self.perf_obj.min_best_case_uplink_2_4),
                    'max_best_case_uplink_2_4_cisco9120' : self._notation_helper(self.perf_obj.max_best_case_uplink_2_4),
                    'avg_best_case_uplink_2_4_cisco9120' : self._notation_helper(self.perf_obj.avg_best_case_uplink_2_4),
                    
                    'min_worst_case_uplink_2_4_cisco9120' : self._notation_helper(self.perf_obj.min_worst_case_uplink_2_4),
                    'max_worst_case_uplink_2_4_cisco9120' : self._notation_helper(self.perf_obj.max_worst_case_uplink_2_4),
                    'avg_worst_case_uplink_2_4_cisco9120' : self._notation_helper(self.perf_obj.avg_worst_case_uplink_2_4),

                    # Schreibt die Datendurchsatzwerte im Downlink in dem 2,4 GHz Frequenzband in das Word Template (Kapitel Performance Analyse)
                    'min_best_case_downlink_2_4_cisco9120' : self._notation_helper(self.perf_obj.min_best_case_downlink_2_4),
                    'max_best_case_downlink_2_4_cisco9120' : self._notation_helper(self.perf_obj.max_best_case_downlink_2_4),
                    'avg_best_case_downlink_2_4_cisco9120' : self._notation_helper(self.perf_obj.avg_best_case_downlink_2_4),
                    
                    'min_worst_case_downlink_2_4_cisco9120' : self._notation_helper(self.perf_obj.min_worst_case_downlink_2_4),
                    'max_worst_case_downlink_2_4_cisco9120' : self._notation_helper(self.perf_obj.max_worst_case_downlink_2_4),
                    'avg_worst_case_downlink_2_4_cisco9120' : self._notation_helper(self.perf_obj.avg_worst_case_downlink_2_4),

                    # Schreibt die Datendurchsatzwerte im Uplink in dem 5 GHz Frequenzband in das Word Template (Kapitel Performance Analyse)
                    'min_best_case_uplink_5_cisco9120' : self._notation_helper(self.perf_obj.min_best_case_uplink_5),
                    'max_best_case_uplink_5_cisco9120' : self._notation_helper(self.perf_obj.max_best_case_uplink_5),
                    'avg_best_case_uplink_5_cisco9120' : self._notation_helper(self.perf_obj.avg_best_case_uplink_5),
                    
                    'min_worst_case_uplink_5_cisco9120' : self._notation_helper(self.perf_obj.min_worst_case_uplink_5),
                    'max_worst_case_uplink_5_cisco9120' : self._notation_helper(self.perf_obj.max_worst_case_uplink_5),
                    'avg_worst_case_uplink_5_cisco9120' : self._notation_helper(self.perf_obj.avg_worst_case_uplink_5),

                    # Schreibt die Datendurchsatzwerte im Downlink in dem 5 GHz Frequenzband in das Word Template (Kapitel Performance Analyse)
                    'min_best_case_downlink_5_cisco9120' : self._notation_helper(self.perf_obj.min_best_case_downlink_5),
                    'max_best_case_downlink_5_cisco9120' : self._notation_helper(self.perf_obj.max_best_case_downlink_5),
                    'avg_best_case_downlink_5_cisco9120' : self._notation_helper(self.perf_obj.avg_best_case_downlink_5),
                    
                    'min_worst_case_downlink_5_cisco9120' : self._notation_helper(self.perf_obj.min_worst_case_downlink_5),
                    'max_worst_case_downlink_5_cisco9120' : self._notation_helper(self.perf_obj.max_worst_case_downlink_5),
                    'avg_worst_case_downlink_5_cisco9120' : self._notation_helper(self.perf_obj.avg_worst_case_downlink_5),

                                        # Schreibt die Roamingzeit und Paketverlust Werte über beide Frequenzbänder in das Word Template (Kapitel Fazit und Empfehlungen)
                    'avg_roaming_time_2_4_5_cisco9120' : self._notation_helper(self.roam_obj.avg_roaming_time_2_4_5),
                    'avg_packet_loss_2_4_5_cisco9120' : self._notation_helper(self.roam_obj.avg_packet_loss_2_4_5),

                    # Schreibt die Roamingzeiten für das 2,4 GHz Frequenzband in das Word Template (Kapitel Roaming Analyse)
                    'min_roaming_time_2_4_cisco9120' : self._notation_helper(self.roam_obj.min_roaming_time_2_4),
                    'max_roaming_time_2_4_cisco9120' : self._notation_helper(self.roam_obj.max_roaming_time_2_4),
                    'avg_roaming_time_2_4_cisco9120' : self._notation_helper(self.roam_obj.avg_roaming_time_2_4),

                    # Schreibt die Roamingzeiten für das 5 GHz Frequenzband in das Word Template (Kapitel Roaming Analyse)
                    'min_roaming_time_5_cisco9120' : self._notation_helper(self.roam_obj.min_roaming_time_5),
                    'max_roaming_time_5_cisco9120' : self._notation_helper(self.roam_obj.max_roaming_time_5),
                    'avg_roaming_time_5_cisco9120' : self._notation_helper(self.roam_obj.avg_roaming_time_5),

                    # Schreibt die Paketverluste für das 2,4 GHz Frequenzband in das Word Template (Kapitel Roaming Analyse)
                    'min_packet_loss_2_4_cisco9120' : self._notation_helper(self.roam_obj.min_packet_loss_2_4),
                    'max_packet_loss_2_4_cisco9120' : self._notation_helper(self.roam_obj.max_packet_loss_2_4),
                    'avg_packet_loss_2_4_cisco9120' : self._notation_helper(self.roam_obj.avg_packet_loss_2_4),

                    # Schreibt die Paketverluste für das 5 GHz Frequenzband in das Word Template (Kapitel Roaming Analyse)
                    'min_packet_loss_5_cisco9120' : self._notation_helper(self.roam_obj.min_packet_loss_5),
                    'max_packet_loss_5_cisco9120' : self._notation_helper(self.roam_obj.max_packet_loss_5),
                    'avg_packet_loss_5_cisco9120' : self._notation_helper(self.roam_obj.avg_packet_loss_5)
                    }        
        self.doc.render(context)
        self.doc.save("generated_doc.docx")

    # Ändert die englische Notation zur deutschen und überprüft das Zahlen wie 1,0 als 1 dargestellt werden ohne Zahlen wie 1,08 zu verändern.
    def _notation_helper(self, value_to_notate):
        regex_filter = "\d+,0\d"
        notated_value = str(value_to_notate).replace(".", ",")
        if (re.match(regex_filter, notated_value)):
            return notated_value
        else:
            notated_value = str(notated_value).split(".0",1)[0]
            notated_value = str(notated_value).replace(",0", "")
            return notated_value
