---
execute: 
  echo: false
---

# Variablen, Funktionen und Operatoren {#sec-chapter-variablen-fkts-ops}

::: {.callout-warning}
## Work in Progress
:::

## Variablen

Keine Unterscheidung zwischen Konstanten und Variablen.


### Dynamische Bereiche {#sec-dynamisches-feld}

### Benannte Bereiche

::: {.callout-note}
Eine Tabelle ist ein spezieller benannter Bereich.
:::

## Funktionen {#sec-funktionen}

## Operatoren

Excel hat eine Anzahl vordefinierter Operatoren. Diese Operatoren haben keine direkte Entsprechung als Funktion. 

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
