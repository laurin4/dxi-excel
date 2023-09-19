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

Obwohl Excel Vektorähnliche Bereiche kennt, ist es sinnvolle sich Excels Bereiche immer als Matrizen vorzustellen. 

- Ein Bereich mit nur einer Zelle ist eine 1x1- oder 0-dimensionale Matrix.
- Ein Bereich mit nur einer Zeile oder einer Spalte ist eine 1xn- oder 1-dimensionale Matrix.
- Ein Bereich mit mehreren Zeilen und Spalten ist eine nxm- oder 2-dimensionale Matrix.

Funktionen können Bereiche als Ergebnis haben. In diesem Fall werden die restlichen Werte zeilen und spaltenweise unterhalb bzw. rechts von der entsprechenden Formel ausgegeben

## Tabellen

Tabellen sind benannte Bereiche, die Vektoren enthalten. Tabellen bestehen aus Spalten, in denen die Werte als Vektoren stehen. Excel Tabellen haben immer Überschriften.

Excel kennt zwei Arten von Tabellen: Wertetabellen und Pivot-Tabellen. 

::: {#def-wertetabelle}
**Wertetabellen** sind Listen, die Listen mit gleicher Länge schachteln.
:::

Wertetabellen entsprechen ungefähr einem *Datenrahmen* (engl. *data-frame*) in anderen Programmiersprachen. Der Unterschied zu einem Datenrahmen ist, dass Excels Wertetabellen *keine Vektoren* erzwingen. Deshalb können in der gleichen Spalte einer Wertetabelle verschiedene Datentypen gemischt werden.

Pivot-Tabellen sind Strukturen zur einfachen Datenanalyse. Pivot-Tabellen entsprechen einer tabellarischen Darstellung der Daten. Ihre Funktion ist auf wenige analytische Funktionen beschränkt und die dargestellten Werte lassen sich 

## Adressierung von Datenstrukturen

::: {#def-bezug}
Ein **Bezug** ist 
:::

### Arbeitsblattadressierung

#### Relative Adressen

#### Absolute Adressen

### Tabellenadressierung

Spalten und einzelne Werte können über die Tabellenadressierung abgefragt werden. Das Ergebnis einer solchen Adressierung ist immer ein *dynamisches Feld* bzw. ein *dynamischer Bereich* (s. [Abschnitt @sec-dynamisches-feld]).
