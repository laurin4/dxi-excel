---
execute: 
  echo: false
---

# Variablen, Funktionen und Operatoren {#sec-chapter-variablen-fkts-ops}

::: {.callout-warning}
## Work in Progress
:::

## Variablen

Excel kennt keine strikte Unterscheidung zwischen Konstanten und Variablen. Grundsätzlich sind alle Werte in Excel Konstanten, weil direkt eingegebene oder durch eine Formel erzeugte Werte nicht durch andere Formel verändert werden können.

::: {.callout-warning}
In Excel lassen sich Zellen und Bereiche *sperren*. Damit lassen sich die Werte wie Konstanten behandeln. Diese Sperre ist Teil der *Formatierung* einer Zelle und muss für das gesamte Arbeitsblatt festgelegt werden. Die Sperre kann nicht für einzelne Zellen oder Bereiche festgelegt werden.

Die Sperre eines Arbeitsblatts wird im Menuband `Überprüfen` mit dem Kommando `Blatt schützen` aktiviert. Richten Sie bei diesem Kommando  **kein** Passwort ein.
:::

Ist ein Arbeitsblatt gesperrt, sind die Werte auf diesem Arbeitsblatt unveränderlich und verhalten sich wie Konstanten. Ist ein Arbeitsblatt nicht gesperrt, dann verhalten sich die Werte wie Variablen.

### Dynamische Bereiche {#sec-dynamisches-feld}

Formeln können mehr als einen Wert verarbeiten und mehr als einen Ergebnis liefern. Solche Formeln müssen nur in die linke obere Ecke eines Bereichs eingetragen werden. Excel erkennt automatisch, dass die Formel auf einen Bereich angewendet werden soll und erzeugt die entsprechende Formel für alle Zellen des Bereichs. Das Ergebnis einer solchen Formel ist ein *dynamischer Bereich*.

::: {#def-vektorisieren}
**Vektorisieren** heisst das Erzeugen eines dynamischen Bereichs aus einem statischen Bereich.
:::

::: {#exm-vektorisieren}
## Vektorisieren eines statischen Bereichs

```
=A1:A10
```
::: 

Im Gegensatz zu einem normalen Bereich, ist bei einem dynamischen Bereich nur die linke obere Zelle bekannt. Um trotzdem alle Zellen eines solchen Bereichs zu adressieren, wird die Gatter- (`#`) Notation verwendet.

Das @exm-vektorisieren erzeugt einen Bereich mit 10 Zellen. Die Formel wird in die linke obere Zelle des Bereichs eingetragen. Die Formel wird z.B. in die Zelle `B1` eingetragen. Anschliessend können die Werte im Bereich `B1:B10` über die Gatter-Notation addressiert werden: `B1#`.

Der Vorteil des Vektorisierens ist, dass der Bereich durch das Einfügen neuer Zeilen vergrössert werden kann, ohne dass die nachfolgenden vektorisierten Formeln angepasst werden müssen.

::: {.callout-tip}
## Praxis

Weil Tabellen automatisch vektorisiert werden, ist es einfacher Werte in einer Tabelle zu erfassen bzw. als Tabelle zu importieren (s. @sec-chapter-daten-importieren) und anschliessend über die Tabellenadressierung auf die Werte zu verweisen.
:::

### Benannte Bereiche

Im Funktionsbalken wird ganz links die Adresse der aktuellen Zelle angezeigt. Wenn in dieses Feld geklickt wird, dann kann der Adresse ein Name zugewiesen werden. Auf diese Weise kann eine Zelle oder ein Bereich benannt werden. Der Name eines Bereichs darf nur einmal in einer Arbeitsmappe existieren. Dafür kann in jeder Zelle der Arbeitsmappe dieser Namen als Adresse verwendet werden. Dazu ist es nicht notwendig, den die genaue Arbeitsblatt- und Zellenadresse zu kennen. 

Verweist ein benannter Bereich mit einer Zelle auf den Anfang eines dynamischen Bereichs, dann kann der gesamte dynamische Bereich mit der Gatter-Notation referenziert werden. 

Auf diese Weise lassen sich oft adressierte Zelleen oder Bereiche benennen. Dadurch kann die Adresse für Verwendung in Formeln abstrahiert werden. 

Die Verwendung von benannten Bereichen wird weiter vereinfacht, dass bei der Eingabe von Formeln die Namen der benannten Bereiche als Vorschlag angezeigt werden.

::: {#exm-dynamischer-benannter-bereich}
## Adressierung eines dynamischen benannten Bereichs
```
=umsatz#
```
::: 

Benannte Zellen helfen, auf wichtige Zellen in einer Arbeitsmappe zuzugreifen. 

::: {.callout-note}
## Merke
Ein benannter Bereich entspricht einer Variablen in anderen Programmiersprachen.
:::

::: {.callout-important}
Eine benannte Zelle oder ein benannter Bereich sind immer absolut referenziert.
:::

::: {.callout-note}
Wenn einem benannten Bereich Zellen hinzugefügt werden, dann wird der benannte Bereich um diese Zellen erweitert.
:::

::: {.callout-note}
Eine Tabelle ist ein spezieller benannter Bereich, der als Ganzes oder in Teilen referenziert werden kann (s. [Abschnitt @sec-tabellenadressen]).
:::

## Funktionen {#sec-funktionen}

## Operatoren

Excel hat eine Anzahl vordefinierter Operatoren. Nicht alle Excel Operatoren haben keine direkte Entsprechung als Funktion. 

### Substitution über Funktionspfade

### Substitution mit `LET()`

## Funktionsketten

- Funktionsketten mit LET

### Leere Zellen in einer Funktionskette {#sec-funktionsketten-leerezelle}

Normalerweise werden leere Zellen als Ergebnis einer Funktion durch `0` ersetzt. Dieses Konvertierung findet erst bei der Darstellung des Ergebnisses statt. Innerhalb einer Funktionskette werden leere Zellen als leere Zellen weitergereicht, solange keine Aggregation vorgenommen wird. Es ist deshalb möglich in einer Funktionskette eine Entscheidung mit `ISTLEER()` für den Fall einer leeren Zelle zu treffen.

## Funktionen selbst definieren

LAMBDA Funktion

Strategie der Funktionsentwicklung
- Funktion als Funktionspfad
- Funktion als Funktionskette
- Funktion als LAMBDA Funktion
- Funktion als benanntes Element
