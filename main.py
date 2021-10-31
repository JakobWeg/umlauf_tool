
import pandas as pd

import numpy as np

def Streckendefeinlesen(NamedateiS):    #Funktion zum Einlesen des Umlaufplans (Dateiname wird Übergeben; Rückgabe: Streckenabschnitt und Umlaufrichtung)

    df = pd.read_excel(NamedateiS, sheet_name=['Streckenabschnitte', 'Umlauf_TEST'])   #Einlesen der Exceldatei mit 2 Sheets (Streckenabschnitte und Umlauf)
    #print(df)
    Umlauf = df['Umlauf_TEST']
    Streckenstruktur = df['Streckenabschnitte']
    #print(type(Umlauf))
    #print(Umlauf.iloc[:, [0]])
    AnzStreckenabschnitte = len(Umlauf)

    Umlauf_list = Umlauf["Nr."].tolist()


    Umlauf_Richtung = Umlauf["Richtung\nV = Vorwärts\nR = Rückwärts"].tolist() #Schreiben der Fahrtrichtung für die Streckenabschnitte im Umlauf

    Dateinamen = Streckenstruktur["Name der .xlsx"].tolist()

    Dateinamenselekt = []

    i=0
    while i < AnzStreckenabschnitte:

        Dateinamenselekt.append(Dateinamen[Umlauf_list[i]-1]) #Dateinameselekt füllen mit der Nummer des Streckenabschnitts im Umlauf

        i += 1

    #print(Dateinamenselekt)
    return Dateinamenselekt, Umlauf_Richtung

#Funktion zum Einlesen der Streckendaten für einen Streckenabschnitt
def Streckeeinlesen(Namedatei, index):

    current_line = 0
    skip_line = 1
    i=0

    header = []
    matrix = pd.read_excel(Namedatei)
    #Falls der Zug die Strecke rückwärts befaehret wird die Liste umgekehrt
    if Streckendef[index][1] == 'R':
        print('Reverse')
        AnzZeilen = len(matrix)

        #for line in reversed(matrix):
        #    matrix_rev.append(line)

        matrix_rev = matrix.iloc[::-1]

        neigung = -(matrix_rev.iloc[:, 7])
        matrix_rev.iloc[:, 7] = neigung


        return matrix_rev  # ,header

    else:
        return matrix  # , header


#Funktion zum ABspeichern der finalen CSV-Datei
def Dateispeichern(Ergebnis):
    list_rows = Ergebnis
    np.savetxt("Daten_DB1.xlsx", list_rows, delimiter=";", fmt='%s')  # Zeilen Trenner: ; und Format: String
    df = pd.DataFrame(list_rows)
    df.to_excel('Umlauf_Test.xlsx', index=False, header=False)

if __name__ == '__main__':

    Streckendef = np.transpose(
        Streckendefeinlesen("Umlaufplanung.xlsx"))  # Transponieren für Zeilenweisen Zugriff

    AnzStreckenabsch = len(Streckendef)
    Streckendaten = pd.DataFrame()
    # Streckendaten = []
    i = 0
    # Einlesen der Strackendaten über Schleife und Abspeichern der verschiedenen Abschnitte in Liste:
    while i < AnzStreckenabsch:
        Streckeninfo = Streckeeinlesen(Streckendef[i][0], i)
        Streckendaten = pd.concat([Streckendaten, Streckeninfo], ignore_index=True)

        i += 1
    print(Streckendaten)
    #header = Streckeninfo[1][0]  # In Streckeninfo[1][0] stehen die Spaltenbeschriftungen



    #Ergebnis = np.vstack((header, Finallist))  # Spaltenbeschriftung hinzufügen



    AnzLines = len(Streckendaten)
    # Hier werden die Zeilen, welche doppelt vorhanden sind gelöscht
    i = 1
    j = 0
    linestodelete = []
    while i < AnzLines - 1:
       if Streckendaten.iloc[i][3] == '[BF]' and Streckendaten.iloc[i - 1][3] == '[BF]':
            linestodelete.append(i - 1)
            j += 1
       i += 1

    Anzlinesdelete = len(linestodelete)
    print('Anzahl der gelöschten Zeilen:', Anzlinesdelete)
    i = 0
    j = 0
    #while i < Anzlinesdelete:
        #del Streckendaten[linestodelete[i] - j]  # -j, da die Anzahl der bereits gelöschten Zeilen abgezogen werden muss, damit die Angabe aus Linestodelete noch stimmt
    Streckendaten = Streckendaten.drop(labels=linestodelete, axis=0)
    print('deleted lines:', linestodelete)
    print(j)

    # # Hier muss noch der Weg neu eingetragen werden. Für Berechnungen müssen die Strings ersteinmal in Float umgewandelt werden.
    AnzZeilen = len(Streckendaten)

    # Löschen von ungebrauchten Spalten
    # Streckendaten.drop(Streckendaten.columns[[0, 12, 13, 15, 19, 20, 22, 23, 24]], axis = 1, inplace = True)

    # Weg neu berechnen
    i=0
    anz_zeilen=len(Streckendaten)
    while i < anz_zeilen-1:
        Streckendaten.iloc[i+1, 9] = Streckendaten.iloc[i, 9]+Streckendaten.iloc[i, 8]
        i += 1

    print(type(Streckendaten))

    # Aufruf der Speichern Funktion um die endgültige Datei im .csv Format zu sichern.
    Dateispeichern(Streckendaten)

