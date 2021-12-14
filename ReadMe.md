# Öffentlicher Kartenvorverkauf Skript
## Welche Funktion hat das Skript?
Das Skript kann automatisch Bestellungen verarbeiten, die über das Bestellformular der [HM-Musik](https://www.hm.edu/konzertkarten) Website getätigt wurden. 
Das Funktioniert, indem das Skript automatisch auf das Postfach *XXX@hm.edu* zugreift.
Dabei werden Emails mit dem Anhang ***FormData_FormData.csv*** automatisch heruntergeladen, verarbeitet, und anschließend in den ***öffentlich*** Ordner im Posteinfang verschoben.

Die eigenen Login-Daten werden nach der Installation einmalig in der Konfigurationsdatei **credentials.yaml** eingetragen. Der Nutzername fängt dabei mit ***hm-*** an.
Weiteres dazu folgt in der Installations Anleitung.

## Installation
### Installation Python
Zur Nutzung des Skriptes ist das Installieren von Python erforderlich
* Python in der Verison 3.8.10 kann auf der Python Website hier für [Windows](https://www.python.org/ftp/python/3.8.10/python-3.8.10-amd64.exe) und hier für [MacOSx](https://www.python.org/ftp/python/3.8.10/python-3.8.10-macosx10.9.pkg) heruntergeladen werden.
* Falls die Links nicht funktionnieren sollten, einfach nach der entsprechenden Verison Suchen und diese herunterladen
* Bei der Installation sollte im ersten Dialog unten der Punkt *"Add Python X.X to PATH"* ausgewählt werden und anschließend auf *"Install Now"* geklickt werden. Falls weitere Dialoge angezeigt werden kann einfach alles auf Standardeinstellungen weiter geklickt werden.

* Anschließend kann das Programm von der HM-Musik Cloud heruntergeladen werden, und in einen geeigneten Ordner entpackt werden. z.B. im Dokumente Ordner.
* Auf den Ordner kann unter Windows jetzt ein *Shift+Rechtsklick" gemacht werden und auf "PowerShell/Eingabeaufforderung hier öffnen" klicken.
* Unter OSx kann auf den Ordner ebenfalls per Rechtsklick über *Services* *"new Terminal at Folder"* asugewählt werden
*  Dort muss dann der folgende Befehl eingegeben werden:
***pip install -r requirements.txt***
* Damit werden alle nötigen Packages installiert, was eine Weile dauern kann.
* Während der Zeit kann nun die im *conf* Ordner hinterlgete *credentials.yaml* modifiziert werden, indem diese mit einem Texteditor der Wahl geöffnet wird
* An den angedeuteten Stellen müssen dann die Logindaten zum eigenen HM Account eingetragen werden. Der Nutzername fängt mit **hm-** an.
* Als nächstes muss im config Ordner noch die Datei best_best.txt mit den aktuellen Daten wie Datum und Uhrzeiten zum Einlass aktualisiert werden und gespeichert werden.
* Wenn die Packages installiert wurden kann das Skript mit dem Befehl ***python okkv_v2.py*** über die PowerShell/Eingabeaufforderung gestartet werden
* Um den automatischen E-Mail Versandt zu aktivieren muss der Befehl ***python okkv_v2.py --send_mail*** ausgeführt werden.
* Im Anschluss kann den Anweisungen des Programmes gefolgt werden. Beim ersten mal Ausführen werden dort einmalig ein paar Daten abgefragt.
