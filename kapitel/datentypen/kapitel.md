---
execute: 
  echo: false
---

# Datentypen {#sec-chapter-datentypen}

## Fundamentale Datentypen

Excel kennt sechs fundamentale Datentypen.

- Zahlen
- Zeichenketten
- Wahrheitswerte
- Fehlerwerte
- Formeln
- leere Zelle

### Zahlen

Zahlen werden in Excel immer als Gleitkommazahlen behandelt. Excel kennt von keine ganzen Zahlen. Wenn eine Zahl in eine ganze Zahl umgewandelt wird (z.B. durch die Funktion `GANZZAHL`), wird lediglich der Nachkommaanteil der Zahl auf `0` gesetzt.

Manche Excel Funktionen arbeiten nur mit ganzen Zahlen. In diesen Fällen wird der Nachkommaanteil der Zahl automatisch abgeschnitten.

Durch die Verwendung von Gleitkommazahlen können in Excel Zahlen mit einer Genauigkeit von 15 signifikanten Stellen dargestellt werden. Werden zwei Zahlen addiert, und das Ergebnis mehr als 15 signifikante Stellen hätte, werden die alle Stellen ab der 15. signifikanten Stelle abgeschnitten. Excel versucht diese Fehler möglichst zu vermeiden.

Der Datentyp *Zahlen* wird in Excel mit der Funktion `ISTZAHL()` geprüft. Die Funktion `ISTZAHL()` liefert `WAHR`, wenn der Wert ein Zahl ist, und sonst `FALSCH`.

### Zeichenketten

Zeichenketten heissen in Excel **Text**. Zeichenketten werden von Excel automatisch gewählt, wenn kein anderer Datentyp für eine Eingabe erkannt wurde. Um die automatische Erkennung zu verhindern, muss der Apostroph-Dekorator (s. [Abschnitt @sec-text-dekorator]) verwendet werden.

Im *Formelmodus* müssen Zeichenketten in doppelte Anführungszeichen eingeschlossen werden. Im Wertemodus müssen Zeichenketten meist nicht besonders markiert werden.

Der Datentyp *Zeichenkette* wird in Excel mit der Funktion `ISTTEXT()` geprüft. Die Funktion `ISTTEXT()` liefert `WAHR`, wenn der Wert eine Zeichenkette ist, und sonst `FALSCH`.

Besondere Zeichenketten sind Adressen bzw. im Excel Jargon **Bezüge**. Die Funktion `ISTBEZUG()` prüft, ob eine Zeichenkette eine gültige Excel Adresse ist. Die Funktion `ISTBEZUG()` liefert `WAHR`, wenn die Zeichenkette eine Adresse ist, und sonst `FALSCH`. 

### Wahrheitswerte

Wahrheitswerte heissen in Excel `WAHR` und `FALSCH`. Wahrheitswerte heissen im Excel Jargon **logische Werte**. Während der Werteeingabe werden Wahrheitswerte unabhängig von der Gross- und Kleinschreibung automatisch erkannt. 

Im Formelmodus dürfen Wahrheitswerte *nicht* in Anführungszeichen eingeschlossen werden, weil sonst die Symbole als Zeichenkette behandelt werden. 

Der Datentyp *Wahrheitswerte* wird in Excel mit der Funktion `ISTLOG()` geprüft. Die Funktion `ISTZLOG()` liefert `WAHR`, wenn der Wert ein Wahrheitswert ist, und sonst `FALSCH`.

### Fehlerwerte

Fehlerwerte ist in Excel ein besonderer Datentyp, um Fehler zu signalisieren. Fehler beginnen immer mit einem Gatter (`#`), das von einem sog. *Fehlerbezeichner* gefolgt wird. In der Regel werden Fehlerwerte automatisch erzeugt, es ist aber möglich, Fehlerwerte auch manuell einzugeben. Bei der manuellen Eingabe von Fehlerwerten werden nur die gültigen Fehlerbezeichner akzeptiert. Andere Symbolfolgen werden als Zeichenketten interpretiert.

Excels gültige Fehlerwerte sind `#NV`, `#NULL!`, `#WERT!`, `#BEZUG!`, `#DIV/0!`, `#ZAHL!`, `#NAME?`, `#KALK!`, `#ÜBERLAUF!`, `#DATEN_ABRUFEN!` und `#ZAHL!`. 

Der Wert `#NV` wird in Excel verwendet, um einen ungültigen Wert zu verweisen. Dieser Wert ist nicht gleichbedeutend mit fehlenden Werten. 

Der Datentyp *Fehlerwert* wird in Excel mit der Funktion `ISTFEHLER()` geprüft. Die Funktion `ISTFEHLER()` liefert `WAHR`, wenn der Wert ein Fehlerwert ist, und sonst `FALSCH`. Um einen bestimmten Fehlerwert zu prüfen, kann die Funktion `FEHLER.TYP()` mit dem entsprechenden Fehlerbezeichner als Argument verwendet werden. Diese Funktion gibt einen eindeutigen Zahlenwert für den Fehler zurück. Ist der Wert kein Fehlerwert, dann wird der Fehler `#NV` zurückgegeben.

### Formeln

Formeln sind ein eigener Datentyp, die mit dem Gleich-Dekorator (`=`) beginnen. Eine Formel besteht immer aus einem Wert, einem Funktionsaufruf oder einer Kombination von Werten und Funktionsaufrufen, die durch Operatoren verknüpft wurden.

Der Datentyp *Formel* wird in Excel mit der Funktion `ISTFORMEL()` geprüft. Die Funktion `ISTFORMEL()` liefert `WAHR`, wenn der Wert eine Formel ist, und sonst `FALSCH`.

Formeln sind in Excel ein besonderer Datentyp, denn eine Zelle mit einer Formel hat immer zwei Datentypen. Der erste Datentyp ist der Datentyp *Formel*, der zweite Datentyp ist der Datentyp des Formelergebnisses.

### Leere Zellen

Lehre Zellen sind Zellen, die keinen Wert enthalten. Leere Zellen zeigen fehlende Werte in Excel an. 

Die Leere Zelle ist für Excel ein eingener Datentyp und wird mit der Funktion `ISTLEER()` geprüft. Die Funktion `ISTLEER()` liefert `WAHR`, wenn der Wert eine leere Zelle ist, und sonst `FALSCH`.

::: {.callout-warning}
Eine leere Zelle und eine leere Zeichenkette lassen sich mit dem Auge nicht unterscheiden. Die Funktion `ISTLEER()` liefert für eine leere Zeichenkette `FALSCH` zurück, weil die Zelle einen Wert enthält.
:::

Leere Zellen als Ergebnis einer Formel werden von Excel in die Zahl `0` umgewandelt. Im Kapitel [Abschnitt @sec-funktionen] wird gezeigt, wie sich leere Zellen von Zellen mit dem Wert `0` unterscheiden lassen. Im [Abschnitt @sec-funktionsketten-leerezelle] wird auf leere Zellen als Funktionsergebnisse ausführlich eingegangen.

## Dekoratoren

::: {#def-dekorator}
Ein **Dekorator** ist Sprachelement einer Programmiersprache, mit dem die normale Interpretation von Symbolen verändert werden kann.
:::

Excel hat zwei Dekoratoren, als erstes Symbol einer Zelle stehen müssen. D.h. es dürfen auch keine Leerzeichen vor einem Dekorator stehen. Die beiden Dekoratoren sind:

- Der Apostroph-Dekorator (`'`)
- Der Gleich-Dekorator (`=`)

### Apostroph-Dekoraktor {#sec-text-dekorator}

Der Apostroph-Dekorator (`'`) erzwingt, dass die nachfolgenden Symbole als Zeichenkette interpretiert werden müssen. 

Mithilfe des Apostroph-Dekorators wird die automatische Typerkennung für die laufende Eingabe deaktiviert. Auf diese Weise lassen sich Zeichenketten eingeben, die nur aus Ziffern bestehen oder Zeichenketten, die Datentumswerten ähneln.

Der Apostroph-Dekorator muss zwingend verwendet werden, wenn auf eine Zeichenkette eine der folgenden Bedingungen zutrifft.

- Die Zeichenkette beginnt mit einem Gleichheitszeichen (`=`), einem Pluszeichen (`+`), einem Minuszeichen (`-`) oder einem Prozentzeichen (`%`).
- Die Zeichenkette besteht nur aus Ziffern, dem Dezimaltrenner (`.`) oder dem Tausendertrenner (`‘`) enthält
- Die Zeichenkette ähnelt der wissenschaftliche Notation für Zahlen mit und ohne Vorzeichen für den Exponenten. Z.B. `5.E3`.
- Die Zeichenkette ähnelt einm Datum.
- Die Zeichenkette entspricht einen Fehlerwert, unabhängig von der Schreibweise. Z.B. `#nv`.
- Die Zeichenkette entspricht einem Wahrheitswert (`WAHR` oder `FALSCH`) unabhängig von der Schreibweise.

Das Apostroph kann entfallen, wenn die aktive Zelle vorher bereits als Datentyp `Text` formatiert wurde. In diesem Fall verlieren alle Symbole ihre spezielle Bedeutung. In solche Zellen können keine Formeln eingegeben werden.

### Gleich-Dekorator {#sec-formel-dekorator}

Der Gleich-Dekorator (`=`) zeigt an, dass die nachfolgenden Symbole als Formel interpretiert werden müssen.

Bei der Eingabe des Gleich-Dekorators wechselt Excel für die laufende Eingabe in den Formelmodus. Mit beeenden der Eingabe mit der Eingabetaste wird der Formelmodus wieder verlassen und das Ergebnis der eingegebenen Formel wird in der Zelle angezeigt.

## Komplexe Datenstrukturen

Excel kennt zwei komplexe Datenstrukturen: 

- Bereiche
- Tabellen

Excels komplexe Datenstrukturen müssen einen fundamentalen Datentyp haben und können keine komplexen Datenstrukturen schachteln.

### Bereiche: Vektoren und Matrizen

Excels Grundstruktur ist das Rechteck, dass durch die markierten Zellen entsteht. Dieses Rechteck heisst im Excel Jargon ein **Bereich**. 

Obwohl Excel vektorähnliche Bereiche kennt, ist es sinnvolle sich Excels Bereiche immer als Matrizen vorzustellen. 

- Ein Bereich mit nur einer Zelle ist eine 1x1- oder 0-dimensionale Matrix.
- Ein Bereich mit nur einer Zeile oder einer Spalte ist für Excel eine 1xn- oder 1-dimensionale Matrix. Diese speziellen Matrizen werden in der Regel als *Vektoren* bezeichnet.
- Ein Bereich mit mehreren Zeilen und Spalten ist eine nxm- oder 2-dimensionale Matrix.

Funktionen können Bereiche als Ergebnis haben. In diesem Fall werden die restlichen Werte zeilen und spaltenweise unterhalb bzw. rechts von der entsprechenden Formel ausgegeben.

### Tabellen

Tabellen sind benannte Bereiche, die Vektoren enthalten. Tabellen bestehen aus Spalten, in denen die Werte als Vektoren stehen. Excel Tabellen haben immer Überschriften.

Excel kennt zwei Arten von Tabellen: 

- Wertetabellen 
- Pivot-Tabellen. 

::: {#def-wertetabelle}
**Wertetabellen** sind Listen, die Listen mit gleicher Länge schachteln.
:::

**Wertetabellen** entsprechen ungefähr einem *Datenrahmen* (engl. *data-frame*) in anderen Programmiersprachen. Der Unterschied zu einem Datenrahmen ist, dass Excels Wertetabellen *keine Vektoren* erzwingen. Deshalb können in der gleichen Spalte einer Wertetabelle verschiedene Datentypen gemischt werden.

::: {#def-pivottabelle}
**Pivot-Tabellen** sind ein Werkzeug zum *interaktiven Zusammenfassen* von Daten.
:::

**Pivot-Tabellen** erleichtern einfache Datenanalysen. Pivot-Tabellen entsprechen einer tabellarischen Darstellung der Daten. Ihre Funktion ist auf wenige analytische Funktionen beschränkt und die dargestellten Werte lassen sich nur umständlich weiterverarbeiten. 

*Pivot-Tabellen* sind Arbeitsmappenelemente aber ***keine Datenstrukturen***. Deshalb werden Pivot-Tabellen in diesem Buch ähnlich einer Visualisierung als *Darstellungsform* behandelt. 

## Adressierung von Datenstrukturen {#sec-adressierung-ds}

::: {#def-bezug}
Ein **Bezug** ist die Adresse eines Bereichs.
:::

In Excel gibt es zwei Arten von Bezügen:

- Arbeitsblattadressen
- Tabellenadressen

### Arbeitsblattadressen

::: {#def-arbeitsblattadresse}
Eine **Arbeitsblattadresse** enthält die Koordinaten eines Bereichs auf einem Arbeitsblatt.
:::

Arbeitsblattadressen besteht aus drei Teilen:

- Arbeitsblattname
- Bereichsbeginn
- Bereichsende

Der Bereichsbeginn verweist immer auf die linke obere Zelle des Bereichs. Der Bereichsende verweist immer auf die rechte untere Zelle des Bereichs. Eine Zelle wird immer durch den Spaltenindex (Buchstabe) und den Zeilenindex (Zahl) identifiziert. Werden die Koordinaten des Bereichsbeginns und des Bereichsendes bei der Eingabe vertauscht, dann wird der Bereich automatisch korrigiert. 

::: {#exm-bereichsadresse}
## Arbeitsblattadresse
```
Tabelle1!A1:C3
```
:::

@exm-bereichsadresse verweist auf den Bereich mit drei Spalten und drei Zeilen auf dem Arbeitsblatt `Tabelle1` beginnend mit der Zelle `A1` und endend mit der Zelle `C3`. 

Oft werden Arbeitsblattadressen nicht vollständig sondern gekürzt angegeben. Es gibt zwei Möglichkeiten, um Arbeitsblattadressen zu kürzen:

- Werden Bereiche auf dem gleichen Arbeitsblatt adressiert, dann kann der Tabellenname weggelassen werden. 
- Wird ein Bereich mit nur einer Zelle adressiert, dann wird das Bereichsende weggelassen.

Weil Arbeitsblattadressen von vielen interaktiven Excelkommandos verwendet werden, gibt es zwei Arten von Arbeitsblattadressen:

- Relative Adressen
- Absolute Adressen

Die Art der Adresse legt fest, wie ein interaktives Kommando mit einer Adresse umgehen soll. Die populärste interaktive Funktion ist das *Autoauffüllen*. Dabei wird eine Zelle mit einer Formel interaktiv auf einen Bereich von Zellen übertragen.

::: {.callout-warning}

Das Autoauffüllen ist eine einfache und beliebte Methode, um Formeln in Excel auf verschiedene Werte wiederholt anzuwenden. Bis 2019 war das Autoauffüllen die einzige Möglichkeit für die Datentransformation.

Die relative und absolute Adressierung ist eine wichtige Voraussetzung für das Autoauffüllen.  Leider ist das Autoauffüllen auch die Ursache für viele Fehler beim Umgang mit Excel.
:::

::: {.callout-tip}
## Praxis

In Excel365 kann das Autoauffüllen durch vektorisierte Funktionen fast vollständig ersetzt werden. Dadurch lassen sich viele Excel-typische Fehler vermeiden. Dadurch ist die Unterscheidung zwischen der relativen und absoluten Adressierung nicht mehr so wichtig.
:::
#### Relative Adressen

::: {#def-relative-adresse}
Eine **relative Adresse** ist eine Adresse eines Bereichs, die *veränderlich* ist.
:::

Relative Adressen werden in Excel von interaktiven Funktionen, wie dem *Autoauffüllen* verwendet, um die Adressen automatisch anzupassen. Eine relative Adresse wird relativ zur aktuellen Zelle angepasst. 

Ein Beispiel für eine relative Adresse ist `A1`. Diese Adresse bezeichnet die Zelle, die sich in der ersten Zeile und der ersten Spalte auf dem aktuellen Arbeitsblatt befindet. Wird die adressierte Zelle interaktiv nach unten aufgefüllt, dann wird die Adresse automatisch zu `A2`, `A3`, usw. angepasst. Wird die adressierte Zelle nach rechts aufgefüllt, dann wird die Adresse automatisch zu `B1`, `C1`, usw. angepasst.

#### Absolute Adressen

::: {#def-absolute-adresse}
Eine **absolute Adresse** ist eine Adresse eines Bereichs, die ganz oder teilweise *unveränderlich* ist.
:::

Der unveränderliche Teil einer Arbeitsblattadresse wird mit einem Dollarzeichen (`$`) markiert. Dieser Teil der Adresse wird bei der Anpassung der Adresse nicht verändert. So lassen sich Adressen angeben, die durch interaktive Kommandos nicht verändert werden.

Ein Beispiel für eine absolute Adresse ist `$A$1`. Diese Adresse bezeichnet die Zelle, die sich in der ersten Zeile und der ersten Spalte auf dem aktuellen Arbeitsblatt befindet. Wird die adressierte Zelle interaktiv horizontal oder vertikal aufgefüllt, dann wird die Adresse nicht angepasst.

Auf diese Weise lassen sich konstante Werte in Formeln einbauen.

### Tabellenadressen

Spalten und einzelne Werte können über die Tabellenadressierung abgefragt werden [@microsoft_support_using_2023]. Das Ergebnis einer solchen Adressierung ist immer ein *dynamisches Feld* bzw. ein *dynamischer Bereich* (s. [Abschnitt @sec-dynamisches-feld]).

::: {.callout-note}
Jede Tabellenadresse kann auch als Arbeitsblattadresse dargestellt werden. Umgekehrt ist dies nicht möglich.
:::

Eine Tabellenadresse besteht aus zwei Teilen:

- Dem Tabellennamen
- Dem Spaltennamen

::: {#exm-tabellenadresse}
## Tabellenadresse
```
Tabelle1[Spalte1]
```
:::

Das @exm-tabellenadresse verweist auf die Spalte `Spalte1` der Tabelle `Tabelle1`. 

::: {.callout-important}
Die Namen von Tabellen sind *unabhängig* den Arbeitsblattnamen. In der gleichen Arbeitsmappe darf jeweils nur eine Tabelle und nur ein Arbeitsblatt mit dem gleichen Namen existieren. Es ist aber möglich, dass eine Tabelle und ein Arbeitsblatt den gleichen Namen haben. Das kann zu Verwirrungen führen, denn die Adresse `Tabelle1[Spalte1]` und die Adresse `Tabelle1!A:A` verweisen nicht zwingend auf die gleiche Spalte, denn eine Tabelle muss nicht auf einem Arbeitsblatt mit dem gleichen Namen stehen!
:::

Tabellenadressen können auch gekürzt werden: 

- Werden Spalten in der gleichen Tabelle angesprochen, dann kann der Tabellenname weggelassen werden.
- Soll die gesamte Tabelle angesprochen werden, dann kann der Spaltenname weggelassen werden.

#### Tabellenbereiche 

Um mehrere Spalten einer Tabelle anzusprechen kann der Bereichsoperator (`:`) wie bei der Arbeitsblattadressierung verwendet werden. Zusätzlich müssen die Spalten in ein weiteres Paar eckiger Klammern eingeschlossen werden. 

::: {#exm-tabellenbereich}
## Tabellenbereich

```
Tabelle1[[Spalte1]:[Spalte3]]
```
:::

@exm-tabellenbereich verweist auf alle Spalten zwischen `Spalte1` und `Spalte3` der Tabelle `Tabelle1`. 

::: {.callout-warning}
Werden zu einem späteren Zeitpunkt neue Spalten zwischen den adressierten Spalten hinzugefügt, dann werden die neuen Spalten automatisch in die Adressierung einbezogen.
:::

::: {.callout-important}
Es ist nicht möglich mehrere Spalten einer Tabellen *gleichzeitig* gezielt zu adressieren, wenn diese nicht unmittelbar nebeneinander stehen. Diese Adressierung ist unmöglich, weil Excel nur rechteckige Bereiche adressieren kann und solche Spalten keinen *rechteckigen Bereich* bilden.

Die Adressierung aus @exm-tabellenbereich kann nicht angepasst werden, so dass nur `Spalte1` und `Spalte3` adressiert werden, ausser die Spalten `Spalte1` und `Spalte3` stehen direkt nebeneinander.
:::

#### Überschriften adressieren

Eine Excel Tabelle hat immer Spaltenüberschriften. Diese Überschriften können ebenfalls über die Tabellenadressierung adressiert werden. Dazu wird als Spaltenname der Wert `#Kopfzeilen` verwendet. 

::: {#exm-kopfzeilen}
## Kopfzeilen einer Tabelle adressieren

```
Tabelle1[#Kopfzeilen]
```
:::

Um gezielt Kofzeilen zu adressieren, kann die Adressierung aus @exm-kopfzeilen mit der Adressierung aus @exm-tabellenbereich kombiniert werden (@exm-kopfzeilen-bereich).

::: {#exm-kopfzeilen-bereich}
## Kopfzeilen einer Tabelle adressieren

```
Tabelle1[[#Kopfzeilen];[Spalte1]:[Spalte3]]
``` 
:::

#### Absolute Adressierung in Tabellen

Ein Tabellenbereich ist immer absolut adressiert. Wird aber nur eine einzelne Spalte einer Tabelle adressiert, dann wird die Adresse relativ addressiert. Damit beim Autoauffüllen diese Adresse nicht verändert wird, muss die Adresse als Bereich angegeben werden.

::: {#exm-absoluter-tabellenbereich-eine-spalte}
## Absolute Tabellenadressierung von einer Spalte

```
Tabelle1[[Spalte1]:[Spalte1]]
```
:::

Wird eine Formel mit dieser Adresse interaktiv aufgefüllt, dann wird die Adresse nicht verändert.
