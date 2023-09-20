---
execute: 
  echo: false
---

# Daten importieren {#sec-chapter-daten-importieren}


::: {.callout-warning}
## Work in Progress
:::

Viele Excel Arbeitsmappen kombinieren Berechnungen und Daten. Excel kann allerdings auch Daten aus anderen Quellen importieren. Quellen können andere Dateien, Datenbanken oder Web-APIs sein und in verschiedenen Formaten vorliegen. Dafür stellt Excel Parser für verschiedene Dateiformate bereit, damit die Daten importiert werden können. 

Daten werden korrekt mit dem Kommando **Daten abrufen (Power Query)** importiert. Das Kommando ist in der Gruppe **Daten** im Abschnitt **Daten abrufen und transformieren** zu finden. Das Kommando *Daten abrufen* startet die sog. **Power Query** Umgebung. In dieser Umgebung können Daten aus verschiedenen Quellen importiert und vor der Bereitstellung korrigiert werden. 

::: {.callout-tip}
# Power Query

**Power Query** ist eine Umgebung für den Datenimport, die von verschiedenen Microsoft Produkten verwendet wird. Die Umgebung ist in Excel, Power BI, Power Apps und Power Automate verfügbar. Power Query stellt eine einheitliche Schnittstelle für den Datenimport bereit und basiert auf einer *Import-Beschreibungssprache*. Mit diese Sprache lassen sich Daten für die Arbeit vorbereiten, so dass viele Datenbereinigungsschritte in Excel Arbeitsmappen entfallen können.

Zu den zentralen Funktionen von Power Query gehören:

- Quellenmanagement
- Überschriftenerkennung
- Schemaerkennung und -transformation
- Entfernen von Duplikaten und ungültigen Werten
- Vektorisierung von Daten
- Kombinieren von Daten aus verschiedenen Quellen
:::

::: {.callout-warning}
Viele tabellarische Dateiformate lassen sich direkt mit Excel öffnen. Das sollte nur mit Excel Arbeitsmappen erfolgen. Bei anderen Dateiformaten kann das direkte öffnen zu Datenverlusten oder Datenfehlern führen. Diese lassen sich in Excel nur umständlich korrigieren. 

Ausserdem lassen sich direkt geöffnete Dateien nicht mehr erweitern, was die Datenerhebung erschwert.
:::

Das Kernprinzip von *Power Query* ist das Verbinden einer Datenquelle mit einer Arbeitsmappen.  Die Datenquelle wird dabei nicht in die Arbeitsmappe übernommen, sondern nur eine Verbindung zu der Datenquelle hergestellt. Dadurch kann eine Datenquelle in mehreren Arbeitsmappen verwendet werden und sich ändern, ohne dass die Arbeitsmappe angepasst werden muss. 

::: {.callout-note}
## Merke
Das Ergebnis eines Imports ist immer eine *Tabelle*.
:::

## Eine Datenverbindung herstellen

1. Daten abrufen
2. Datenquelle auswählen
3. Daten überprüfen
4. Daten gegebenenfalls transformieren
5. Daten importieren

Nach einem Datenimport liegen die Daten in der Arbeitsmappe vor. Diese Daten sind eine Kopie der Daten in der Datenquelle. Die kopierten Werte sind als Tabelle in der Arbeitsmappe gespeichert. Dadurch kann die Arbeitsmappe unabhängig von der Datenquelle verwendet und geteilt werden.

::: {.callout-warning}
## Achtung

Excel betrachtet jeden Datenimport als Sicherheitsproblem. Daher werden Datenverbindungen beim Öffnen einer Arbeitsmappe standardmässig deaktiviert. Beim Öffnen einer Arbeitsmappe mit einer Datenverbindung wird eine Warnung angezeigt. Um mit den importierten Daten arbeiten zu können, muss die Datenverbindung *aktiviert* werden.

Wird die Datenverbindung nicht aktiviert, werden Funktionen die auf die Daten zugreifen nicht ausgeführt, sondern nur die Ergebnisse der letzten Ausführung angezeigt.
:::


## Daten aktualisieren

Die Daten können sich aber ändern, z.B. weil ein Formular ausgefüllt wurde oder eine Datenbank aktualisiert wurde. Dadurch ändern sich die Daten in der Datenquelle. In solchen Fällen ist die Kopie in der Arbeitsmappe nicht mehr aktuell. Um die Daten zu aktualisieren, muss die Datenverbindung aktualisiert werden.

Dazu wird das Kommando **Alle aktualisieren** verwendet. Das Kommando ist im Menu **Daten** im Abschnitt **Verbindungen** zu finden. Das Kommando aktualisiert alle Datenverbindungen in der Arbeitsmappe entsprechend der Importspezifikation.

ABBILDUNG ALLE AKTUALISIEREN

Es werden nur die importierten Daten aktualisiert. Wurde die Tabelle durch Formeln erweitert, werden diese Formeln *nicht* gelöscht, sondern auf die neuen Daten erweitert. 

::: {.callout-important}
Die importierte Struktur darf durch Formeln **nicht verändert** werden. Das bedeutet, dass keine Spalten innerhalb der Struktur hinzugefügt oder gelöscht werden dürfen. Neue Spalten für Formeln müssen rechts von der importierten Datenstruktur der Tabelle hinzugefügt werden.
:::

## Datenverbindung anpassen

Gelegentlich ändert sich der Ort einer Datenquelle. Sehr häufig passiert das, wenn eine Datei in einen anderen Ordner verschoben wird. In solchen Fällen muss die Datenverbindung angepasst werden. Dazu muss in Power Query die Datenquelle angepasst werden.

VORGEHENSWEISE

1. Schritt "Quelle" ausählen
2. Pfad zur Datenquelle anpassen
3. Schema überprüfen
4. Daten aktualisieren

Alle Datenverbindungen werden auch im Dialog **Abfragen & Verbindungen** angezeigt. Der Dialog ist in der Gruppe **Verbindungen** im Abschnitt **Daten** zu finden. Der Dialog zeigt alle Datenverbindungen an, die in der Arbeitsmappe verwendet werden. 

ABBILDUNG VERBINDUNGEN ANZEIGEN

::: {.callout-warning}
## Veraltete Praxis
Grundsätzlich lassen sich die Datenquellen auch im Dialog **Abfragen & Verbindungen** anpassen. Diese Anpassungen können jedoch nicht kontrolliert werden und die Datenstruktur der neuen Quelle wird von Excel direkt übernommen.

ABBILDUNG VERBINDUNG ANPASSEN

Wird bei diesem Schritt eine falsche Quelle ausgewählt, verändert sich möglicherweise die importierte Datenstruktur und **die Formeln in der importierenden Arbeitsmappe werden wegen falscher Adressierungen zerstört**.

Dieses Problem besteht bei der Anpassung von Datenquellen in Power Query nicht, weil die zu importierende Datenstruktur vor dem Import gegen das urspüngliche Datenschema geprüft wird. Gravierende Änderungen werden angezeigt und die Anpassung kann abgebrochen oder korrigiert werden.
::: 

## Daten beim Import transformieren {#sec-daten-import-transformieren}

### Datenschema anpassen

#### Schritt Höherstufen

#### Schritt Geänderter Spaltentyp

#### Schritt Spalten entfernen

### Daten bereinigen

### Datenquellen kombinieren
