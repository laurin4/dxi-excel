---
# bibliography: references.bib

abstract: ""

execute: 
  echo: false
---

# Zeichenketten {#sec-chapter-zeichenketten}

::: {.callout-warning}
## Work in Progress
:::


## Zeichenketten trennen


### Festkodierte Werte trennen

Eine festkodierte Datenstruktur ist eine Zeichenkette, die Werte an festgelegten Positionen mit konstanten Längen enthält. Diese Daten lassen sich mit der Funktion `TEIL()` extrahieren. 

Die Funktion hat drei Argumente:

- Die Zeichenkette, aus der die Daten extrahiert werden sollen.
- Die Position, an der die Daten beginnen.
- Die Länge der Daten als Anzahl von Symbolen.

::: {#exm-iban-trennen}
## IBAN in Land, Prüfziffer, Bankkennung und Kontonummer trennen

Die IBAN ist eine festkodierte Datenstruktur. Die IBAN enthält die Länderkennung, die Prüfziffer, die Bankkennung und die Kontonummer. Die Länge der einzelnen Daten ist konstant und die Position der Felder ist festgelegt. 

| Feld | Position | Länge |
|:------|---:|---:|
| Land | 1 | 2 |
| Prüfziffer | 3 | 2 |
| Bankkennung | 5 | 5 |
| Kontonummer | 10 | Länge der IBAN - 10 |

Die (ungültige) IBAN `CH12BANK1002135135` kann mit der `TEIL()`-Funktion in die einzelnen Felder zerlegt werden. Dazu müssen zuerst die Positionen und Längen der Felder erstellt werden. Dazu werden die Positionen und Längen der Felder untereinander geschrieben.

| A | B | C | D |
| --- | --- | --- | --- |
| 1 | 3| 5 | 10 |
| 2 | 2 | 5 | `= LÄNGE(IBAN_Nummer) - 10` |

Diese Werte werden als Vektoren der Funktion `TEIL()` übergeben. 

```
= TEIL(IBAN_Nummer; A1:D1; A2:D2)
```

Die Funktion `TEIL()` gibt die einzelnen Felder als Vektor zurück. Das Ergebnis ist `{"CH"; "12"; "BANK1"; "002135135"}`. Hier muss berücksichtigt werden, dass die einzelnen Felder weiterhin Zeichenketten sind.
:::

### Einzelne Symbole extrahieren

Eine Zeichenkette kann als eine festkodierte Datenstruktur verstanden werden. Jedes Symbol in der Zeichenkette hat eine feste Position und eine feste Länge `1`. Diese Position ist identisch mit dem Index des Symbols in der Zeichenkette. Weil die Positionen der Symbole in einer Zeichenkette eine Sequenz ist, können alle Positionen mithilfe der Länge der Zeichenkette generiert werden.

> ::: {#exm-symbole-trennen}
> ## Symbole einer Zeichenkette in Adresse `A1` trennen
> 
> ```
> =TEIL(A1;SEQUENZ(LÄNGE(A1));1)
> ```
> 
> oder als horizontaler Vektor
> 
> ```
> =TEIL(A1;SEQUENZ(1;LÄNGE(A1));1)
> ```
> :::
