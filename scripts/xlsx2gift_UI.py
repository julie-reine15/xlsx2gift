import os
from re import S
import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox
)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("XLSX to GIFT Converter")
        self.setGeometry(450, 100, 700, 200)
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.label = QLabel("Sélectionnez un fichier XLSX à convertir.")
        self.layout.addWidget(self.label)

        self.select_xlsx_btn = QPushButton("Choisir le fichier XLSX")
        self.select_xlsx_btn.clicked.connect(self.select_xlsx)
        self.layout.addWidget(self.select_xlsx_btn)

        self.selected_xlsx_label = QLabel("")
        self.layout.addWidget(self.selected_xlsx_label)

        self.save_txt_btn = QPushButton("Choisir l'emplacement et le nom du fichier TXT de sortie")
        self.save_txt_btn.clicked.connect(self.save_txt)
        self.save_txt_btn.setEnabled(False)
        self.layout.addWidget(self.save_txt_btn)

        self.selected_txt_label = QLabel("")
        self.layout.addWidget(self.selected_txt_label)

        self.convert_btn = QPushButton("Convertir")
        self.convert_btn.clicked.connect(self.convert)
        self.convert_btn.setEnabled(False)
        self.layout.addWidget(self.convert_btn)

        self.xlsx_path = ""
        self.txt_path = ""

    def select_xlsx(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel Files (*.xlsx)")
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                self.xlsx_path = selected_files[0]
                self.selected_xlsx_label.setText(f"Fichier XLSX sélectionné : {self.xlsx_path}")
                self.save_txt_btn.setEnabled(True)

    def save_txt(self):
        file_dialog = QFileDialog(self)
        file_dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        file_dialog.setNameFilter("Text Files (*.txt)")
        file_dialog.setDefaultSuffix("txt")
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                self.txt_path = selected_files[0]
                self.selected_txt_label.setText(f"Fichier TXT de sortie : {self.txt_path}")
                self.convert_btn.setEnabled(True)

    def convert(self):
        try:
            # print(self.txt_path)
            print("Converting...")
            data_p = Data_processor(self.xlsx_path, self.txt_path)
            data_p.xlsx_to_db()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Une erreur est survenue : {e}")
        # try:
        #     data_p = Data_processor(self.xlsx_path)
        #     # Patch: pass output path to write_into_file
        #     def write_file_patch(content):
        #         data_p.write_into_file(content, self.txt_path)
        #         return None
        #     # Patch the last cont to write to the selected file
        #     data_p.xlsx_to_db(cont=lambda db: data_p.split_db_by_question_type(db, cont=lambda _: None))
        #     QMessageBox.information(self, "Succès", "Conversion terminée !")
        # except Exception as e:
        #     QMessageBox.critical(self, "Erreur", f"Une erreur est survenue : {e}")

class Data_processor:
    def __init__(self, file_path: str, output_path: str) -> None:
        self.file_path = file_path
        self.output_path = output_path

    def xlsx_to_db(self, cont = lambda x:x):
        return cont(self.clean_db(pd.read_excel(self.file_path, 1)))

    def clean_db(self, question_db: pd.DataFrame, cont = lambda x:x):
        return cont(self.split_db_by_question_type(question_db.dropna(subset=[question_db.columns[0]])))

    def split_db_by_question_type(self, cleaned_db: pd.DataFrame, cont = lambda x:x):
        question_types = cleaned_db.drop_duplicates(subset=cleaned_db.columns[0], ignore_index=True)["code_type_question"].to_list()
        for type in question_types:
            cont(self.db_to_gift_syntax(cleaned_db[cleaned_db["code_type_question"]==type], int(type)))

    def db_to_gift_syntax(self, db_by_question_type: pd.DataFrame, type_q: int, cont = lambda x:x, output_path=None):
        gift_format_questions = ""
        l_strip = lambda x : x.strip(" ")
        # print(db_by_question_type["intitule_question"])
        
        db_by_question_type = db_by_question_type.copy() # to avoid the SettingWithCopyWarning 
        db_by_question_type.loc[:, "code_question"] = db_by_question_type["intitule_question"].apply(lambda x: f"::{x}::")

        match type_q:
            case 1: # vrai/faux
                # pass
                # code question - énoncé question - vrai/faux - feedback
                db_by_question_type["vrai_faux"] = db_by_question_type["vrai_faux"].apply(lambda x :str(x).upper())
                for _, question in db_by_question_type[["code_question", "enonce_question", "vrai_faux", "feedback"]].iterrows():
                    gift_format_questions += (str(question["code_question"])+
                                                question["enonce_question"]+
                                                "{"+question["vrai_faux"])
                    if isinstance(question["feedback"], str):
                        gift_format_questions += "####" + question["feedback"] + "}\n\n"
                    else:
                        gift_format_questions += "}\n\n"

            case 2: # choix multiple
                # pass
                # code question - énoncé question - réponse correcte - réponse incorrecte - réponse multiple - coefficient - feedback
                for _, question in db_by_question_type[["code_question", "enonce_question", "reponse_correcte", "reponse_incorrecte", "reponse_multiple", "coefficient", "feedback"]].iterrows():
                    if not question["reponse_multiple"]:
                        # print(question[4])
                        gift_format_questions += str(question["code_question"]+
                                                    question["enonce_question"]+
                                                    " {\n"+
                                                    "\t="+
                                                    "\n\t=".join(map(l_strip, question["reponse_correcte"].split(";"))) +
                                                    "\n\t~"+
                                                    "\n\t~".join(map(l_strip, question["reponse_incorrecte"].split(";")))
                                                    )
                        if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                        else:
                                gift_format_questions += "}\n\n"

                    else:
                        # print("~".join(map(l, question[3].split(";"))))
                        gift_format_questions += str(question["code_question"]+
                                                    question["enonce_question"]+
                                                    " {\n"+
                                                    "\t~"+
                                                    "\n\t~".join([f"%{b}%{a}" for a,b in zip(map(l_strip, question["reponse_correcte"].split(";")), map(l_strip, question["coefficient"].split(";")))])+
                                                    "\n\t~"+
                                                    "\n\t~".join(map(l_strip, question["reponse_incorrecte"].split(";"))))
                        if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                        else:
                                gift_format_questions += "}\n\n"
                # print(gift_format_questions)
            case 3: # appariement
                # pass
                # code question - énoncé question - réponse correcte - feedback
                l_appariement = lambda x: "\t="+x.replace("=", "->").strip()
                for _, question in db_by_question_type[["code_question", "enonce_question", "reponse_correcte", "feedback"]].iterrows():
                    # print(question[2].replace("=", "->").replace(";", "\n\t"))
                    gift_format_questions += str(question["code_question"]+
                                                question["enonce_question"]+
                                                " {\n" + 
                                                "\n".join(map(l_appariement, question["reponse_correcte"].split(";"))))
                    if isinstance(question["feedback"], str):
                        gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                    else:
                        gift_format_questions += "}\n\n"
            
            case 5: # numérique
                # pass
                for _, question in db_by_question_type[["code_question", "enonce_question", "reponse_correcte", "marge_erreur", "reponse_multiple", "coefficient", "feedback"]].iterrows():
                    # print(type(question["feedback"]))
                    gift_format_questions += str(question["code_question"]+
                                                question["enonce_question"]+
                                                " {#"
                                            )
                    if not question["reponse_multiple"]: # si une seule bonne réponse
                        if isinstance(question["reponse_correcte"], str): # si la réponse est une intervalle
                            gift_format_questions += question["reponse_correcte"].replace("_", "..") 
                            if isinstance(question["feedback"], str):
                                gift_format_questions += " ####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                        if isinstance(question["reponse_correcte"], int) or isinstance(question["reponse_correcte"], float) : # si la réponse est un chiffre
                            if question["marge_erreur"] > 0:
                                gift_format_questions += str(question["reponse_correcte"]) + ":" + str(question["marge_erreur"])
                            else:
                                gift_format_questions += str(question["reponse_correcte"])
                            if isinstance(question["feedback"], str):
                                    gift_format_questions += "\n####" + question['feedback']+ "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                    else: # si plusieurs réponses possibles
                        if isinstance(question["coefficient"], str) and isinstance(question["marge_erreur"], str): # coefficient et marge d'erreur pour toutes les réponses
                            gift_format_questions += "\n\t="
                            gift_format_questions += "\n\t=".join([f"%{a}%{b}:{c}" for a,b,c in zip(map(l_strip, question["coefficient"].split(";")), map(l_strip, question["reponse_correcte"].split(";")), map(l_strip, question["marge_erreur"].split(";")))])
                            if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                        elif isinstance(question["coefficient"], str) and isinstance(question["marge_erreur"], float): # coefficient pour toutes les réponses mais pas de marge d'erreur
                            gift_format_questions += "\n\t="
                            gift_format_questions += "\n\t=".join([f"%{a}%{b}" for a,b in zip(map(l_strip, question["coefficient"].split(";")), map(l_strip, question["reponse_correcte"].split(";")))])
                            if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                        elif isinstance(question["coefficient"], float) and isinstance(question["marge_erreur"], str): # coefficient pour toutes les réponses mais pas de marge d'erreur
                            gift_format_questions += "\n\t="
                            gift_format_questions += "\n\t=".join([f"{a}:{b}" for a,b in zip(map(l_strip, question["reponse_correcte"].split(";")), map(l_strip, question["marge_erreur"].split(";")))])
                            if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                        else:
                            gift_format_questions += "\n\t="
                            gift_format_questions += "\n\t=".join([a for a in map(l_strip, question["reponse_correcte"].split(";"))])
                            if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                                
            case 4: # réponse courte
                # pass
                for _, question in db_by_question_type[["code_question", "enonce_question", "reponse_correcte", "feedback"]].iterrows():
                    gift_format_questions += str(question["code_question"]+
                                                question["enonce_question"]+
                                                " {\n"+
                                                "\t="+
                                                "\n\t=".join(map(l_strip, question["reponse_correcte"].split(";"))))
                    if isinstance(question["feedback"], str):
                        gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                    else:
                        gift_format_questions += "}\n\n"
            case 6: # mot manquant
                # pass
                # TODO: modify the code to handle only one missing word
                for _, question in db_by_question_type[["code_question", "enonce_question", "reponse_correcte", "reponse_incorrecte", "feedback"]].iterrows():
                    nb_trou = str(question["enonce_question"]).count("{}")
                    final_gift_question = ""
                    for i in range(nb_trou):
                        answers_set = ""
                        for j in range(len(question["reponse_correcte"].split(";"))):
                            if j==i :
                                answers_set += "=" + question["reponse_correcte"].split(";")[j].strip(" ") + " "
                            else:
                                answers_set += "~" + question["reponse_correcte"].split(";")[j].strip(" ") + " "
                        if isinstance(question["reponse_incorrecte"], str):
                            answers_set += "\n~" + "\n~".join(map(l_strip, question["reponse_incorrecte"].split(";"))) + ""
                            
                        question["enonce_question"] = str(question["enonce_question"]).replace("{}", "{"+answers_set+"}", 1)
                        final_gift_question = question["enonce_question"]
                    
                    
                    gift_format_questions += question["code_question"] + final_gift_question
                    
                    if isinstance(question["feedback"], str):
                        gift_format_questions += " ####" + question["feedback"] + "\n\n"
                    else :
                        gift_format_questions += "\n\n"
                        
        # print(gift_format_questions)
        cont(self.write_into_file(gift_format_questions))

    def write_into_file(self, content: str):
        
        if os.path.getsize(self.output_path) > 0:
            with open(self.output_path, "a", encoding="utf-8") as file:
                file.write(content)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())