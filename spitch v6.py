# -*- coding: utf-8 -*-
"""
Created on Tue Dec 15 14:29:18 2020

@author: User
"""

import os
import pandas as pd
from datetime import datetime
import win32com.client
from win32com.client import Dispatch, constants
from jinja2 import FileSystemLoader, Environment
import numpy as np
from matplotlib import pyplot as plt
import seaborn as sns
import glob
from numpy import random


path = "C:/Users/User/Desktop/Python Scripting/10 spitch/"
os.chdir(path)


# df
df = pd.read_excel("spitch.xlsx", sheet_name="S - 20,21 - Daten")
# ref_spieler = df.pivot_table(index = "Spieler", values = "Punkte", aggfunc=np.mean).to_excel("ref_spieler.xlsx")
ref_spieler = pd.read_excel("ref.xlsx")
df_zitate =  pd.read_excel("ref.xlsx",sheet_name = "zitate")
df_header_pics =  pd.read_excel("ref.xlsx",sheet_name = "header_pics")
# benötigt zur pivotisierung - pivot_table kann eine col nicht als values und cols nehmen
df["Platzierung Spieltag c"] = df["Platzierung Spieltag"]
# random zitat und header_pic
# int from min
a = datetime.now()
a = int(a.strftime("%S")) + int(a.strftime("%M"))
np.random.seed(a)
# np.random.seed(np.random.randint(1000))
# zitat
# randomisierte zahl aus allen zitatindex
zitat_index = np.random.randint(1,len(df_zitate))
# auswahl des einen zitats über den index
zitat = df_zitate.at[zitat_index, "zitate"]
# hinzufügen des zitatlinks zum ref_framework
ref_spieler["zitat"] = zitat
# header_pic - gleiches prozedere wie beim zitat
header_pic_index = np.random.randint(1,len(df_header_pics))
header_pic = df_header_pics.at[header_pic_index,"header_pics"]
ref_spieler["header_pic"] = header_pic




print("import")
# stats letzter Spieltag
letzter_spieltag = df[df.Spieltag == df.Spieltag.max()]

# identifikation der Spieler

##### winner of the day
# funktion zur berechnung der gebrauchten Variablen
def wotd_vars (Spieltag = letzter_spieltag): 
    df_wotd = letzter_spieltag[letzter_spieltag.Punkte == letzter_spieltag.Punkte.max()]
    wotd_spieler = df_wotd["Spieler"].iloc[0]
    wotd_platzierung_gesamt = df_wotd["Platzierung Gesamt"].iloc[0]
    wotd_pic_link = df_wotd.merge(ref_spieler[["Spieler","pic"]], on = "Spieler").iloc[0]["pic"]
    wotd_punkte = df_wotd["Punkte"].iloc[0]
    return df_wotd, wotd_spieler, wotd_platzierung_gesamt, wotd_pic_link, wotd_punkte

df_wotd, wotd_spieler, wotd_platzierung_gesamt, wotd_pic_link, wotd_punkte = wotd_vars()


##### loser of the day
# funktion zur berechnung der gebrauchten Variablen
def lotd_vars (Spieltag = letzter_spieltag): 
    df_lotd = letzter_spieltag[letzter_spieltag.Punkte == letzter_spieltag.Punkte.min()]
    lotd_spieler = df_lotd["Spieler"].iloc[0]
    lotd_platzierung_gesamt = df_lotd["Platzierung Gesamt"].iloc[0]
    lotd_pic_link = df_lotd.merge(ref_spieler[["Spieler","pic"]], on = "Spieler").iloc[0]["pic"]
    lotd_punkte = df_lotd["Punkte"].iloc[0]

    return df_lotd, lotd_spieler, lotd_platzierung_gesamt, lotd_pic_link, lotd_punkte

df_lotd, lotd_spieler, lotd_platzierung_gesamt, lotd_pic_link, lotd_punkte = lotd_vars()


##### goat
# identifikation des goat
ident_goat = df.groupby("Spieler").sum().reset_index()
# ident_goat["Rang_Schnitt"] = ident_goat["Platzierung Spieltag"] / df["Spieltag"].max()
df_goat = ident_goat[ident_goat.Punkte == ident_goat.Punkte.max()]


####### alternative ideen - goat index
# ident_gooat = pd.pivot_table(df,index="Platzierung Spieltag", 
#                         columns="Spieler", 
#                         values = "Platzierung Spieltag c", 
#                         aggfunc="count")

# platzierung = ident_gooat.index

# df_test = ident_gooat.mul(platzierung, axis=0)
# gooat = df_test.mean().sort_values()


def goat_vars():
    goat_sum = ident_goat[ident_goat.Punkte == ident_goat.Punkte.max()]
    goat = goat_sum["Spieler"].iloc[0]
    goat_punkte_gesamt = df_goat["Punkte"].iloc[0]
    goat_filter = df["Spieler"] == goat
    goat_df = df[goat_filter]
    goat_punkte_schnitt = goat_df["Punkte"].mean().round(2)
    goat_punkte_max = goat_df["Punkte"].max()
    goat_filter2 = ref_spieler["Spieler"] == goat
    goat_pic_link = ref_spieler[goat_filter2]["pic"].iloc[0]
    return goat, goat_punkte_gesamt, goat_punkte_schnitt, goat_punkte_max, goat_pic_link
  
goat, goat_punkte_gesamt, goat_punkte_schnitt, goat_punkte_max, goat_pic_link = goat_vars()  




# df_goat_index = df_platzierungen.iloc[0,:].sort_values(ascending=False).to_frame()
# df_goat_index = df_goat_index.add_prefix("Anzahl Platz ").reset_index()
# # df_goat_index["P pro Platz"] = np.arange(1,0,-0.125)
# df_ppp = pd.Series(np.arange(1,0,-0.125), name = "ppp").to_frame()
# df_goat_index = df_goat_index.merge(df_ppp, left_index = True, right_index = True)
# pivot = df_goat_index.pivot_table(index = "Anzahl Platz 1", values = "ppp", aggfunc="mean").reset_index()
# df_goat_index = df_goat_index.merge(pivot, how = "outer", on = "Anzahl Platz 1")




##### spender

ident_spender = df.pivot_table(index = "Spieler", values = "Platzierung Spieltag c", columns="Platzierung Spieltag", aggfunc = "count").reset_index()
ident_spender.columns.astype(str)
ident_spender = ident_spender.add_prefix("Platz")
df_spender = ident_spender[ident_spender.Platz8 == ident_spender.Platz8.max()]
df_spender = df_spender.rename(columns= {"PlatzSpieler" :"Spieler"})

def spender_vars():
    spender = df_spender["Spieler"].iloc[0]
    spender_verlorene_spieltage = df_spender["Platz8"].iloc[0]
    spender_verlorene_spieltage = int(spender_verlorene_spieltage)
    spender_betrag = spender_verlorene_spieltage * 5
    spender_pic_link = df_spender.merge(ref_spieler[["Spieler","pic"]], on = "Spieler").iloc[0]["pic"]

    # spender_pic_link
    return spender, spender_verlorene_spieltage, spender_betrag, spender_pic_link

spender, spender_verlorene_spieltage, spender_betrag, spender_pic_link = spender_vars()

#### sieger der herzen

sdh_pic_link = "https://i.ibb.co/mtkHdwj/felix-mack-1024x1024.jpg"



#### Differenz des Spieltags
null_filter = df["Punkte"] != 0
df_spieltage_ohne_null = df[null_filter]

df_spieltage = df_spieltage_ohne_null.pivot_table(index = "Spieltag", values="Punkte", aggfunc=["min","mean", "max"])
df_spieltage["Differenz"] = df_spieltage["max"] - df_spieltage["min"]


### differenz zwischen Letztem und Vorletztem
platzierung = [7,8]
# loserS of the day
df_lsotd_filter = df["Platzierung Spieltag"].isin(platzierung)
df_lsotd = df[df_lsotd_filter]
df_lsotd_pivot = df_lsotd.pivot(index = "Spieltag", columns = "Platzierung Spieltag", values = "Punkte")
df_lsotd_pivot = df_lsotd_pivot.add_prefix("Platz: ")
df_lsotd_pivot["Kellerduell"] = df_lsotd_pivot["Platz: 7"] - df_lsotd_pivot["Platz: 8"]
kellerduell_diff = df_lsotd_pivot["Kellerduell"].iloc[-1]

### der deckel

deckel_filter = df["Platzierung Spieltag"] == 8
df_letzter = df[deckel_filter]
deckel = df_letzter.pivot_table(index = "Spieler",values = "Platzierung Spieltag", aggfunc = "count")
deckel = deckel.rename(columns = {"Platzierung Spieltag" : "Packungen"})
deckel["Betrag"] = deckel["Packungen"] * 5
deckel["Betrag"] = deckel["Betrag"].astype(str) + "€"

# deckel["Betrag"] = str(deckel["Betrag"]) + "€"
deckel = deckel.sort_values(by = "Packungen", ascending=False)

deckel.index.name = None
deckel_html = deckel.to_html(classes = "", table_id = "my_table")
print(deckel_html)

print("calculations")
############### Grafiken




dpi = 500

label_font_size = 7

# Meine Ballons
g01 = sns.relplot(
    data=df,
    x="Spieltag", y="Punkte",
    hue="Spieler", size="Platzierung Spieltag",
    sizes=(200, 10))
plt.xticks(df["Spieltag"], fontsize = label_font_size)
plt.savefig("Ballons.jpeg", dpi = dpi, bbox_inches = "tight")

# plt.gca().invert_xaxis()

# meine Spiele
g02 = sns.relplot(
    data=df,
    x="Spieler", y="Punkte",
    hue="Spieler",)
# plt.xticks(df["Spieltag"])
g02.set_xticklabels(rotation = 45)
plt.savefig("games.jpeg", dpi = dpi, bbox_inches = "tight")


# def Xgesamt_übersicht ():
    # gesamtübersicht 
    # g1 = sns.scatterplot(data = df, 
    #              x = "Spieltag",
    #              y = "Punkte",
    #              hue = "Spieler")
    # plt.legend(bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0.)
    # plt.xticks(df["Spieltag"])
    # plt.savefig("Gesamtübersicht.jpeg", dpi = 1000, bbox_inches = "tight")


# spielerübersicht in boxplots
plt.clf()
# df_box = pd.pivot(df,index = "Spieltag", columns="Spieler", values = "Punkte" )
# g2 = sns.boxplot(data = df_box)
# g2.set_xticklabels(rotation = 45, labels =  df_box.columns)
# # plt.legend(bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0.)
# plt.savefig("Spielerübersicht.jpeg", dpi = 1000, bbox_inches = "tight")


#### spieltage im zeitverlauf
def spieltage_plot(): 
    g05 = sns.relplot(
    data=df_spieltage,
    x="Spieltag", y="Differenz"
    )
    plt.ylim(0,df_spieltage["Differenz"].max()+300)
    plt.xticks(df["Spieltag"], fontsize = label_font_size)
    plt.title("Differenz zwischen Winner und Loser of the Day")
    plt.tight_layout()
    plt.savefig("Differenz - Spieltage.jpeg", dpi = dpi)

tust = spieltage_plot()

# , bbox_inches = "tight"

# # individuelle Entwicklung im Zeitverlauf
def ind_ent (df = df, beschreibung = " - Entwicklung", pic_format = ".jpeg"):
    path_ind_ent = "C:/Users/User/Desktop/Python Scripting/10 spitch/ind_ent"
    os.chdir(path_ind_ent)
    df = df
    spieler_liste = list(df["Spieler"].unique())
    for element in spieler_liste:
        x = df["Spieler"] == element
        data = df[x]
        g = sns.lmplot(data = data,
                        x="Spieltag",
                        y="Punkte")
        plt.xticks(df["Spieltag"])
        plt.ylim(0,5000)
        plt.title(element)
        savename = element + beschreibung + pic_format   
        plt.savefig(savename,dpi = dpi, bbox_inches = "tight")       
    os.chdir(path)

# ind_ents = ind_ent()

##### 5 Spieltagetrend
trend = df["Spieltag"].max() - 5
df_trend = df[df.Spieltag > trend]

ind_trend = ind_ent(df = df_trend, beschreibung= " 5 Spieltage Trend")


print("pics")

# individuelle änhänge einsammeln
ind_ent_path = "C:/Users/User/Desktop/Python Scripting/10 spitch/ind_ent"
# os.chdir(ind_ent_path)
int_ent_list = glob.glob(ind_ent_path + "/*.jpeg")
ref_spieler["fig_link"] = int_ent_list


# alle allgemeinen anhänge einsammeln
all_jpeg = glob.glob(path + "/*.jpeg")

####################

path_template = "C:/Users/User/Desktop/Python Scripting/10 spitch/email template/regular"
name_template = "spitch template.html"

# wotd_link_describ = ""
# lotd_link_describ = ""

print("start mailing")
 

def df_mailing_with_template (df, action = None, general_at = [], name_spalte = "Name", email_spalte = "email", email_text = ""): 
    # read in df
    df = df
    action = action
    name_spalte = name_spalte
    # open outlook
    const=win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    file_loader = FileSystemLoader(path_template)
    env = Environment(loader=file_loader)
    template = env.get_template(name_template)
    

    for i, series in df.iterrows(): 
        # create new Mail Object
        newMail = obj.CreateItem(olMailItem)
        # format message
        newMail.BodyFormat = 2
        # define target 
        newMail.To = series.loc[email_spalte]
        # define Subject
        newMail.Subject = "Spitch...Spiiiitch....SPIIIIITCH"
        Spieler = series.loc["Spieler"]
        
        # Anrede, AiT, Getränk = series.loc["Anrede"], series.loc["Alter in Tage - Zahl"], series.loc["Getränk"]
        # Getränke_link = series.loc["Link"]
        
        output = template.render(Spieler = Spieler,
                                 #header und zitat
                                 header_pic = header_pic,
                                 zitat = zitat,
                                 # wotd
                                 wotd_spieler = wotd_spieler,
                                 wotd_platzierung_gesamt= wotd_platzierung_gesamt, 
                                 wotd_pic_link = wotd_pic_link,
                                 wotd_punkte = wotd_punkte,
                                 wotd_kommentar = "069BoyXXX ist wieder im Tritt...die Leistungskurve zeigt steil nach oben und er zeigt uns was er von der Konkurrenz hält",
                                 # loser of the day
                                 lotd_spieler = lotd_spieler,
                                 lotd_punkte = lotd_punkte,
                                 lotd_pic_link = lotd_pic_link,
                                 kellerduell_diff = kellerduell_diff,
                                 kellerduell_kommentar = "Ein Kommentar zum Kellerduell macht nur Sinn, wenn es ein Duell gab....aber Wrong Rightson war dieses Wochenende wohl mit seinen Gedanken hauptsächlich auf dem CatWalk unterwegs",
                                 # goat"
                                 goat_pic = "https://media.giphy.com/media/3ohhwAADQcO4Ti5LSo/giphy.gif",
                                 goat_punkte_max = goat_punkte_max,
                                 goat_punkte_schnitt = goat_punkte_schnitt,
                                 goat_punkte_gesamt = goat_punkte_gesamt, 
                                 goat_pic_link = "https://giphy.com/embed/ahZZZZFGLGhvq",
                                 goat_kommentar = "",
                                 # spender
                                 spender_verlorene_spieltage = deckel_html, 
                                 spender_betrag = "",
                                 spender_pic_link = "https://christianritter.files.wordpress.com/2012/06/bierdeckel.jpg?w=833", 
                                 spender_kommentar = "Nächstes Mal ist die Tabelle vielleicht hübscher...vielleicht auch nicht. Aber ich schaue jetzt CL",
                                 kommentar_der_regie = "Ankündigung: Jaja, Spieltag 27 fehlt noch, I know!", 
                                 # sieger derherzen
                                 sdh_kommentar = "Garry, Garry, Garry",
                                 sdh_pic_link = sdh_pic_link
                                 )
        newMail.HTMLBody = output
        
        # Attachments - for all
        if general_at == []:
            pass
        else:
            for element in general_at:
                newMail.Attachments.Add(Source=element)
        
        # attachments individual - also as lists?
        if "fig_link" in df.columns:
            newMail.Attachments.Add(Source = series.loc["fig_link"])
        # action
        if action == "send":
            newMail.send()
        elif action == "display":
            newMail.display()
        else: 
            newMail.save()


just_do_it = df_mailing_with_template(df = ref_spieler, general_at=all_jpeg)

print("fin")

print(deckel_html)
