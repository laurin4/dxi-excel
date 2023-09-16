---
execute: 
  echo: false
---

# Datentypen {#sec-chapter-datentypen}

## Fundamentale Datentypen

- leere Zelle
- Wahrheitswerte
- Zahlen
- Zeichenketten
- Fehlerwerte
- Formeln

## Dekoratoren

Problem Automatische Typumwandlung

Excel kennt 

- Der Tick-Dekorator (`'`)
- Der Gleich Dekorator (`=`)

## Bereiche: Vektoren und Matrizen

Excels Grundstruktur ist das Rechteck, dass durch die markierten Zellen entsteht. Dieses Rechteck heisst im Excel Jargon ein **Bereich**. 

Obwohl Excel Vektorähnliche Bereiche kennt, ist es sinnvolle sich Excels Bereiche immer als Matrizen vorzustellen. 

- Ein Bereich mit nur einer Zelle ist eine 1x1- oder 0-dimensionale Matrix.
- Ein Bereich mit nur einer Zeile oder einer Spalte ist eine 1xn- oder 1-dimensionale Matrix.
- Ein Bereich mit mehreren Zeilen und Spalten ist eine nxm- oder 2-dimensionale Matrix.

Excels Bereiche haben *immer* einen fundamentalen Datentyp. 

## Tabellen

Tabellen sind benannte Bereiche, die Vektoren enthalten. Tabellen bestehen aus Spalten, in denen die Werte als Vektoren stehen. Excel Tabellen haben immer Überschriften.

Excel kennt zwei Arten von Tabellen: Wertetabellen und Pivot-Tabellen. 
