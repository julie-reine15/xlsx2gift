import os
import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox
)
import traceback


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
            tb = traceback.format_exc()
            QMessageBox.critical(self, "Erreur", f"Une erreur est survenue : {e}\n\nTraceback:\n{tb}")
            
            
class Data_processor:
    def __init__(self, file_path: str, output_path: str) -> None:
        self.file_path = file_path
        self.output_path = output_path

    def xlsx_to_db(self, cont = lambda x:x):
        """_summary_
        starts the data processing from xlsx to gift format, transforming the data from the xlsx file to a pandas dataframe.
        Args:
            cont (_type_, optional): _description_. Defaults to lambdax:x.

        Returns:
            _type_: _description_
        """
        return cont(self.clean_db(pd.read_excel(self.file_path, 1))) # 1 to read the second sheet of the excel file, first sheet is the template instructions

    def clean_db(self, question_db: pd.DataFrame, cont = lambda x:x):
        """_summary_
        cleans the dataframe from the xlsx file, removing empty rows and splitting the dataframe by question type.
        Args:
            question_db (pd.DataFrame): a pandas dataframe containing the questions from the xlsx file : all the rows, including empty ones.
            cont (_type_, optional): _description_. Defaults to lambdax:x.

        Returns:
            _type_: _description_
        """
        return cont(self.split_db_by_question_type(question_db.dropna(subset=[question_db.columns[0]]))) # focusing on the first column to remove empty rows, because it is mandatory to fill it in the xlsx template to declare a question

    def split_db_by_question_type(self, cleaned_db: pd.DataFrame, cont = lambda x:x):
        """_summary_
        splits the cleaned dataframe by question type and sends each sub-dataframe to be converted to gift format.
        from now on we don't use return anymore, but continuations (cont) to pass the data from one function to another. otherwise, only the last kind of question would be processed and returned.
        Args:
            cleaned_db (pd.DataFrame): a cleaned pandas dataframe containing the questions from the xlsx file and no empty rows.
            cont (_type_, optional): _description_. Defaults to lambdax:x.
        """
        question_types = cleaned_db.drop_duplicates(subset=cleaned_db.columns[0], ignore_index=True)["code_type_question"].to_list() # getting the different question types in the dataframe
        for type in question_types: # one after the other, the different question types are processed
            cont(self.db_to_gift_syntax(cleaned_db[cleaned_db["code_type_question"]==type], int(type))) # filtering the dataframe by question type and sending it to be converted to gift format

    def db_to_gift_syntax(self, db_by_question_type: pd.DataFrame, type_q: int, cont = lambda x:x):
        """_summary_

        Args:
            db_by_question_type (pd.DataFrame): a pandas dataframe containing the questions from the xlsx file filtered by question type.
            type_q (int): the question type code (1 to 6).
            cont (_type_, optional): _description_. Defaults to lambdax:x.
        """
        gift_format_questions = "" # to store the gift formatted questions
        l_strip = lambda x : x.strip(" ") # to remove leading and trailing spaces from strings
        
        db_by_question_type = db_by_question_type.copy() # to avoid the SettingWithCopyWarning 
        db_by_question_type.loc[:, "code_question"] = db_by_question_type["intitule_question"].apply(lambda x: f"::{x}::") # adding the question title syntax to the question title (done 1 time for all questions, i guess we worry about time complexity lol)

        match type_q:
            case 1: # vrai/faux
                # pass
                db_by_question_type["vrai_faux"] = db_by_question_type["vrai_faux"].apply(lambda x :str(x).upper()) # making sure the true/false answers are in uppercase so it matches GIFT syntax
                for _, question in db_by_question_type[["code_question", "enonce_question", "vrai_faux", "feedback"]].iterrows(): # iterating through the questions of this type and formatting them in GIFT syntax
                    gift_format_questions += (str(question["code_question"])+
                                                question["enonce_question"]+
                                                "{"+question["vrai_faux"])
                    if isinstance(question["feedback"], str): # checking if there is feedback to add it to the question
                        gift_format_questions += "####" + question["feedback"] + "}\n\n"
                    else:
                        gift_format_questions += "}\n\n"

            case 2: # choix multiple
                # pass
                for _, question in db_by_question_type[["code_question", "enonce_question", "reponse_correcte", "reponse_incorrecte", "reponse_multiple", "coefficient", "feedback"]].iterrows():
                    if not question["reponse_multiple"]: # choix multiple questions can have one or multiple correct answers, here we handle the case of one correct answer
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

                    else: # handling the case of multiple correct answers
                        gift_format_questions += str(question["code_question"]+
                                                    question["enonce_question"]+
                                                    " {\n"+
                                                    "\t~"+
                                                    "\n\t~".join([f"%{b}%{a}" for a,b in zip(map(l_strip, question["reponse_correcte"].split(";")), map(l_strip, question["coefficient"].split(";")))])+
                                                    ("\n\t~"+
                                                        "\n\t~".join(map(l_strip, question["reponse_incorrecte"].split(";")))
                                                        if isinstance(question["reponse_incorrecte"], str)
                                                        else ""
                                                    )   
                        )
                        if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                        else:
                                gift_format_questions += "}\n\n"
            case 3: # appariement
                l_appariement = lambda x: "\t="+x.replace("=", "->").strip() # replacing the equal sign used in the xlsx template by the arrow used in GIFT syntax
                for _, question in db_by_question_type[["code_question", "enonce_question", "reponse_correcte", "feedback"]].iterrows():
                    gift_format_questions += str(question["code_question"]+
                                                question["enonce_question"]+
                                                " {\n" + 
                                                "\n".join(map(l_appariement, question["reponse_correcte"].split(";"))))
                    if isinstance(question["feedback"], str):
                        gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                    else:
                        gift_format_questions += "}\n\n"
            
            case 5: # numérique
                for _, question in db_by_question_type[["code_question", "enonce_question", "reponse_correcte", "marge_erreur", "reponse_multiple", "coefficient", "feedback"]].iterrows():
                    gift_format_questions += str(question["code_question"]+
                                                question["enonce_question"]+
                                                " {#"
                                            )
                    if not question["reponse_multiple"]: # if there is only one possible answer
                        if isinstance(question["reponse_correcte"], str): # if the answer is an range of values
                            gift_format_questions += question["reponse_correcte"].replace("_", "..") 
                            if isinstance(question["feedback"], str):
                                gift_format_questions += " ####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                        if isinstance(question["reponse_correcte"], int) or isinstance(question["reponse_correcte"], float) : # if the answer is a single value
                            if question["marge_erreur"] > 0: # if there is a margin of error
                                gift_format_questions += str(question["reponse_correcte"]) + ":" + str(question["marge_erreur"])
                            else:
                                gift_format_questions += str(question["reponse_correcte"])
                            if isinstance(question["feedback"], str):
                                    gift_format_questions += "\n####" + question['feedback']+ "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                    else: # if multiple answers are possible
                        if isinstance(question["coefficient"], str) and isinstance(question["marge_erreur"], str): # coefficient and margin of error for each answer
                            gift_format_questions += "\n\t="
                            gift_format_questions += "\n\t=".join([f"%{a}%{b}:{c}" for a,b,c in zip(map(l_strip, question["coefficient"].split(";")), map(l_strip, question["reponse_correcte"].split(";")), map(l_strip, question["marge_erreur"].split(";")))])
                            if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                        elif isinstance(question["coefficient"], str) and isinstance(question["marge_erreur"], float): # coefficient for each answer but no margin of error
                            gift_format_questions += "\n\t="
                            gift_format_questions += "\n\t=".join([f"%{a}%{b}" for a,b in zip(map(l_strip, question["coefficient"].split(";")), map(l_strip, question["reponse_correcte"].split(";")))])
                            if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                        elif isinstance(question["coefficient"], float) and isinstance(question["marge_erreur"], str): # no coefficient and margin of error for each answer
                            gift_format_questions += "\n\t="
                            gift_format_questions += "\n\t=".join([f"{a}:{b}" for a,b in zip(map(l_strip, question["reponse_correcte"].split(";")), map(l_strip, question["marge_erreur"].split(";")))])
                            if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                        else: # no coefficient and no margin of error
                            gift_format_questions += "\n\t="
                            gift_format_questions += "\n\t=".join([a for a in map(l_strip, question["reponse_correcte"].split(";"))])
                            if isinstance(question["feedback"], str):
                                gift_format_questions += "\n\t####" + question["feedback"] + "}\n\n"
                            else:
                                gift_format_questions += "}\n\n"
                                
            case 4: # réponse courte
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
                # this case is handling only one missing word per question, multiple missing words are not supported by the GIFT syntax
                # we can add multiple distractors (wrong answers) though
                for _, question in db_by_question_type[["code_question", "enonce_question", "reponse_correcte", "reponse_incorrecte", "feedback"]].iterrows():
                    final_gift_question = ""
                    answers_set = ""
                    if isinstance(question["reponse_correcte"], str):
                        answers_set += "=" + question["reponse_correcte"] + " "
                    
                    if isinstance(question["reponse_incorrecte"], str):
                        answers_set += "~" + " ~".join(map(l_strip, question["reponse_incorrecte"].split(";"))) + " "
                            
                    question["enonce_question"] = str(question["enonce_question"]).replace("{}", "{"+answers_set+"}")
                    final_gift_question = question["enonce_question"]
                    
                    gift_format_questions += question["code_question"] + final_gift_question
                    
                    if isinstance(question["feedback"], str):
                        gift_format_questions += " ####" + question["feedback"] + "\n\n"
                    else :
                        gift_format_questions += "\n\n"

        cont(self.write_into_file(gift_format_questions))

    def write_into_file(self, content: str):
        
        if os.path.exists(self.output_path):
            if os.path.getsize(self.output_path) > 0:
                with open(self.output_path, "a", encoding="utf-8") as file:
                    file.write(content)
        else:
            with open(self.output_path, "w", encoding="utf-8") as file:
                file.write(content)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())