# un saut de ligne pour séparer 2 questions

# question à choix multiple 
# mauvaises réponses préfixées par un ~
# bonnes réponses préfixées par un = 
# ex : Who's buried in Grant's tomb?{~Grant ~Jefferson =no one}
# The American holiday of Thanksgiving is celebrated on the {
    # ~second
    # ~third
    # =fourth
    # } Thursday of November.

# variante mot manquant -> Les signes _____ ne sont insérés que si les réponses sont placées avant la ponctuation finale.
# ex : Grant is {~buried =entombed ~living} in Grant's tomb.

# réponse courte
# les réponses sont toutes préfixées d'un signe égal (=), indiquant que toutes sont correctes. Les réponses ne peuvent contenir de tilde.
# Who's buried in Grant's tomb?{=no one =nobody}

# vrai/faux
# La réponse s'écrit {TRUE} ou {FALSE}, ou de façon abrégée {T} ou {F}.
# ex : Grant is buried in Grant's tomb.{F}

# appariement 
# Les paires correspondantes doivent commencer par un signe égal (=) et sont séparées par le symbole -> . Il doit y avoir au moins 3 paires.
# Match the following countries with their corresponding capitals. {
    # =Canada -> Ottawa
    # =Italy  -> Rome
    #=Japan  -> Tokyo
    #=India  -> New Delhi
    # }
# pas de feedbacks, ni de coefficients

# numérique 
# la réponse commence par le signe # 
# elle peut contenir une marge d'erreur écrite immédiatement après la réponse correcte, séparée par un signe : pour une réponse située entre 1.5 et 2.5, on écrira {#2:0.5}
# ex : What is the value of pi (to 3 decimal places)? {#3.1415:0.0005}.
# réponse sous la forme d'une intervalle : {#minimum..maximum}
# ex : Quelle est la valeur de pi (3 décimales) ? {#3.141..3.142}
# Si plusieurs réponses numériques sont utilisées, elles doivent être séparées par un signe égal (=), comme les réponses courtes
# ex : When was Ulysses S. Grant born? {#
    # =1822:0
    # =%50%1822:2} -> utilisation d'un coefficient
    
# commentaires
# servent à documenter les questions, commencent par //
# ils ne seront pas importés dans moodle
# ex : // Subheading: Numerical questions below
    # What's 2 plus 2? {#4}

# nom de question
# Un nom de question peut être indiqué en le plaçant avant la question, entouré par des double deux-points
# ex ::Kanji Origins::Japanese characters originally
    # came from what country? {=China}
    
# feedback 
# Un feedback peut être inclus avec chaque réponse. Il suffit de placer le feedback immédiatement après la réponse, séparé par un dièze (#).
# ex : Grant is buried in Grant's tomb.{FALSE#No one is buried in Grant's tomb.}
# /!\ Pour les questions à choix multiples, le feedback n'est affiché que pour la réponse sélectionnée par l'étudiant. Pour la réponse courte, le feedback est affichée uniquement si l'étudiant a donné la réponse correcte correspondante. Pour les questions Vrai/Faux, le feedback n'apparaît que si la réponse donnée est incorrecte.

# coefficients (en %)
# pour les questions à choix multiples et pour les questions à réponse courte
# inclus en faisant suivre le tilde (pour les questions à choix multiples) ou le signe égal (pour les questions à réponse courte) par le pourcentage désiré, entouré de signes pour-cent (par exemple %50%)
# ex : ::Jesus' hometown::Jesus Christ was from {
    # ~Jerusalem#This was an important city, but the wrong answer.
    # ~%25%Bethlehem#He was born here, but not raised here.
    # ~%50%Galilee#You need to be more specific.
    # =Nazareth#Yes! That's right!}.
    
# réponses multiples 
# On active cette option en donnant aux diverses réponses des coefficients
# ex : What two people are entombed in Grant's tomb? {
    # ~No one
    # ~%50%Grant
    # ~%50%Grant's wife
    # ~Grant's father }

import pandas as pd

class Data_processor:
    def __init__(self, file_path: str) -> None:
        self.file_path = file_path

    def xlsx_to_db(self, cont = lambda x:x):
        """_summary_

        Args:
            file_path (str): path to any xlsx file following the template

        Returns:
            _type_: a pandas DataFrame 
        """
        return cont(self.clean_db(pd.read_excel(self.file_path, 1)))

    
    def clean_db(self, question_db: pd.DataFrame, cont = lambda x:x):
        """_summary_
        delete all rows where there's no question text

        Args:
            question_db (pd.DataFrame): pandas dataframe containing the questions to import
            cont (_type_, optional): for continuity puropose
        """
        
        return cont(self.split_db_by_question_type(question_db.dropna(subset=[question_db.columns[0]])))
    
    
    def split_db_by_question_type(self, cleaned_db: pd.DataFrame, cont = lambda x:x):
        question_types = cleaned_db.drop_duplicates(subset=cleaned_db.columns[0], ignore_index=True)["code_type_question"].to_list()

        for type in question_types:
            # print(isinstance(type,int))
            cont(self.db_to_gift_syntax(cleaned_db[cleaned_db["code_type_question"]==type], int(type)))
    
    
    def db_to_gift_syntax(self, db_by_question_type: pd.DataFrame, type_q: int, cont = lambda x:x):
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
        """_summary_

        Args:
            content (str): content to write into the file
        """
        # TODO: change the file path to a variable
        # TODO: check if the file already exists and delete the content first before writing
        with open("/Users/jbaune/Nextcloud/COMP/formations/MOOC_SCIENCE_OUVERTE/banque_question_MOOC_SO_gift.txt", "a", encoding="utf-8") as file:
            # print(content)
            # print(file.write(content))
            file.write(content)


data_p = Data_processor("/Users/jbaune/Nextcloud/COMP/formations/MOOC_SCIENCE_OUVERTE/questions_MOOC_SO.xlsx")
data_p.xlsx_to_db()
