---
bibliography: references.bib

title: Aussagenlogik

abstract: ""

execute: 
  echo: false
---

## Wahrheitswerte in Excel 

Wie im Kapitel Datentypen bereits erwähnt, kennt Excel den Datentyp der Wahrheitswerte. Diese Wahrheitswerte können entweder den Wert `WAHR` oder `FALSCH` haben. Eine Operation, die Wahrheitswerte ergibt wird als *logischer Ausdruck* bezeichnet. Weil logische Ausdrücke für viele Funktionen und Operationen notwendig sind, wandelt Excel die Werte anderer Datentypen bei Bedarf um. Dabei gelten die folgenden Regeln: 

- `0` und die *leere Zelle* entspricht dem Wert `FALSCH`.
- Alle Zahlen ungleich `0` entstprechen dem Wert `WAHR`.
- Alle Zeichenketten inklusive der *leeren Zeichenkette* entsprechen dem Wert `WAHR`.
- Fehlerwerte bleiben unverändert.

Werden Wahrheitswerte in mathematischen Operationen und Funktionen verwendet, dann konvertiert Excel den Wert `FALSCH` in `0` und den Wert `WAHR` in `1` um. 

Werden Wahrheitswerte als Parameter an Zeichenkettenfunktionen übergeben, dann werden die Wahrheitswerte in die entsprechende Zeichenkette umgewandelt. Aus dem Wert `WAHR` wird also die Zeichenkette `"WAHR"` und aus dem Wert `FALSCH` wird die Zeichenkette `"FALSCH"`. 

> **ACHTUNG:** Wahrheitswerte werden in den verschiedenen Sprachversionen von Excel in der eingestellten Sprache angegeben. Beim Wechsel zwischen verschiedenen Excel-Sprachversionen werden die Wahrheitswerte automatisch korrekt angezeigt. Die Umwandlung in Zeichenketten erfolgt dann in der jeweiligen Sprache. Deshalb sollte die Verwendung von in Zeichenketten konvertierten Wahrheitswerten in nachgelagerten Funktionen vermieden werden.

## Aussagenlogische Operationen

Für Wahrheitswerte existieren spezielle Operationen, um die Regeln der *Aussagenlogik* bzw. der *Boole'schen Algebra* abzubilden. Diese Operationen verknüpfen Wahrheitswerte und haben Wahrheitswerte als Ergebnis. Die vier Grundoperationen `NICHT` ($\lnot$), `UND` ($\land$), `ODER` ($\lor$, "inklusives Oder") und `XODER` ($\oplus$, "entweder-oder") sind in Excel als Funktionen verfügbar.

Die Funktion `NICHT()` wandelt einen Wert in den jeweils den anderen Wahrheitswert um. Falls anstelle eines Wahrheitswerts ein anderer Datentyp übergeben wurde, gelten die oben angegebenen Regeln für die Umwandlung. 

Die Funktionen `UND()`, `ODER()` und `XODER()` sind *Aggregatoren*. Das bedeutet, dass Sie alle Werte in dem angegebenen Bereichen zusammenfassen. Das ist oft nicht das gewünschte Verhalten. Deshalb müssen wir für die Arbeit mit Vektoren auf die Bool'sche Arithmetik zurückgreifen.

Zum Beispiel sollen für die folgenden Werte *paarweise* der logische Ausdruck $a \land b$ ausgewertet werden:

|   A |   B |
|:----:|:----:|
| `WAHR` | `FALSCH` |
| `FALSCH` | `WAHR` |
| `WAHR` | `WAHR` |
| `FALSCH` | `WAHR` |
| `WAHR` | `WAHR` |
| `FALSCH` | `FALSCH` |

Die Formel `= UND(A1:A6; B1:B6)` liefert den Wert `FALSCH` zurück, weil nicht alle Werte im *gesamten Bereich* von `A1:A6` und `B1:B6` gleich `WAHR` sind. Es gibt keine Funktion und keinen eigenen logischen Operator für die paarweise logische Verknüpfung dieser beiden Bereiche. Deshalb werden in Excel oft logische Ausdrücke in der Boole'schen Arithmethik eingesetzt, um nur die Werte aus den gleichen Datensätzen miteinander zu vergleichen. Diese Schreibweise ist immer dann notwendig, wenn logische Ausdrücke sich auf die einzelnen Datensätze beziehen.

Die Formel `= A1:A6 * B1:B6` hat die Werte `{0;0;1;0;1;0}` zum Ergebnis. Um Wahrheitswerte zuerhalten kann noch auf die Ungleichheit mit `0` geprüft werden. Dieser Schritt ist in der Praxis aber selten notwendig. Dazu wird die Formel wie folgt ergänzt: `= (A1:A6 * B1:B6) <> 0`. Das Ergebnis ist nun `{FALSCH; FALSCH; WAHR; FALSCH; WAHR; FALSCH}`.

Weil Excel alle Werte ungleich `0` als `WAHR` interpretiert, können die Operationen `UND()` mit `*` und die Operation `ODER()` mit `+` direkt ersetzt werden. 

Die Operation `XODER()` entspricht der Ungleichheit `<>`. Dabei muss allerdings darauf geachtet werden, das dieser Vergleich als Ersatz für `XODER()` *ausschliesslich* für Wahrheitswerte bzw. `0` und `1` erlaubt ist. Für die oben gezeigten Werte ergibt die Formel `= (A1:A5 <> B1:B5)` die Werte `{WAHR; WAHR; FALSCH; WAHR; FALSCH; FALSCH}`. Verwenden wir allerdings die folgenden Werte, dann liefert die Formel `= (A1:A6 <> B1:B6)` die Werte `{WAHR; WAHR; WAHR; WAHR; WAHR; WAHR}`.

|   A |   B |
|:----:|:----:|
| `1` | `0` |
| `0` | `2` |
| `1` | `2` |
| `0` | `1` |
| `3` | `2` |
| `0` | `0` |

Damit das Richtige Ergebnis erzeugt wird, müssen die Werte in Wahrheitswerte konvertiert werden. Dazu muss die Formel durch Vergleiche ungleich `0` erweitert werden: 

```
= (A1:A6 <> 0) <> (B1:B6 <> 0)`
``` 

Diese Formel liefert die gewünschten Werte `{WAHR; WAHR; FALSCH; WAHR; FALSCH; FALSCH}`.

Die folgende Tabelle zeigt logischen Operatoren und die zugehörigen Terme für die Boole'sche Arithmetik.

| Operator | Boole'sche Operation | Vereinfachter Operator |
| :---: | :---: | :---: | 
| $\lnot$ | `1 - a` | `NICHT(a)` |
| $\land$ | `a * b` | `a * b` |
| $\lor$ | `a + b - a * b` | `a + b` |
| $\oplus$ | `a + b - 2 * a * b` oder `( a - b ) ^ 2` | `(a <> b)` |`

## Vergleiche

Eine besondere Art von logischen Ausdrücken sind *Vergleiche*. Ein Vergleich prüft das Verhältnis zweier Werte zueinander. Excel kennt die üblichen Vergleichsoperatoren, die jeweils einen Wahrheitswert zurückliefern.

In Excel werden die Vergleichsoperatoren wie folgt geschrieben: 

- `>` (grösser als)
- `<` (kleiner als)
- `>=` (grösser oder gleich)
- `<=` (kleiner oder gleich)
- `=` (gleich)
- `<>` (ungleich) 

Excel's Vergleichsoperatoren sind Datentypen sensitiv. Das bedeutet, dass die Operatoren Datentypen vor dem Vergleich *nicht* angleichen. Der Vergleich folgende Vergleich ergibt also immer `FALSCH`.

```
= 3 = "3"
```

Weil die Operationen `*`, `+` und `-` normalerweise vor den Vergleichsoperatoren ausgeführt werden, müssen alle Vergleiche eines logischen Ausdrucks für die Boole'sche Arithmetik in Klammern gesetzt werden.

### Der $\in$-Operator mit XVERWEIS

Mithilfe der Funktion `XVERGLEICH()` kann der $\in$-Operator aus der *Mengenlehre* in Excel für logische Ausdrücke bereitgestellt werden. Mit dieser Funktion `XVERWEIS()` können sowohl der $\in$ als auch der $\notin$-Operator mit `XVERWEIS()` abgebildet werden.

Für die folgenden Beispiele verwenden wir die Werte:

|       A       |      B      |
|:-------------:|:-----------:|
| Suchkriterium | Suchbereich |
|       `4`       |      `1`      |
|       `7`       |      `3`      |
|               |      `4`      |
|               |      `8`      |

-   $\in$-Operator: `XVERWEIS(A2:A3; B2:B5; B2:B5 = B2:B5; FALSCH)`
-   $\notin$-Operator: `XVERWEIS(A2:A3; B2:B5; B2:B5 <> B2:B5; WAHR)`

Der Trick besteht darin, dass die *Rückgabematrix* durch einen Vergleich aus dem Suchbereich erzeugt wird. Dadurch wird die Rückgabematrix mit den gleichen Wahrheitswerten für alle Werte im Suchbereich gefüllt. Die erste Formel ergibt deshalb `{WAHR, FALSCH}` und die zweite Formel `{FALSCH, WAHR}`, weil der Wert `4` im Suchbereich vorkommt und der Wert `7` nicht. Diese Werte können direkt in logischen Ausdrücken verwendet werden.

### Zeichenketten vergleichen

Die Vergleichsoperatoren zeigen die Unterschiede zweier Zeichenketten bezüglich der alphabetischen Sortierung an. Die "kleinste" Zeichenkette ist die leere Zeichenkette. Gross- und Kleinschreibung wird bei Zeichenkettenvergleichen nicht unterschieden.

Weil Excel für Vergleiche die nicht-druckbaren Zeichen mit Ausnahme des Leerzeichens und des Tabulators ignoriert, gibt der Vergleichsoperator `=` `WAHR` auch für Zeichenketten zurück, die unterschiedliche nicht-druckbare Zeichen enthalten. Das gleiche Problem entsteht beim Vergleich von unterschiedlichen Schreibweisen. Um auch diese Unterschiede zu erkennen, müssen wir die Funktion `IDENTISCH()` verwenden. Diese Funktion vergleicht die Zeichenketten Zeichen für Zeichen und liefert nur dann `WAHR` zurück, wenn die Zeichenketten exakt gleich sind.

BEISPIEL

## Entscheidungen

Excel kennt zwei Funktionen für Fallunterscheidungen: `WENN()` und `WENNS()`. Die Funktion `WENN()` ist eine *einfache* Unterscheidung, die Funktion `WENNS()` unterstützt *mehrfache* Unterscheidungen. In anderen Programmiersprachen wird in diesem Zusammenhang auch von *Verzweigungen* gesprochen.

### WENN



### WENNS



Mit der Funktion `WENNS()` lassen sich verschachtelte Entscheidungen vereinfachen. 



`WENNS()` kann ausschliesslich logische Ausdrücke mit ihren Ergebnissen verbinden. Anders als bei `WENN()` gibt es keine Möglichkeit, ein Ergebnis festzulegen, falls kein logischer Ausdruck `WAHR` ergibt. Um ein solches Verhalten zu erzeugen, wird ausgenutzt, dass die Funktion `WENNS()` immer einen wahren logischen Ausdruck mit einem Ergebnis verknüpft. Weil die logischen Ausdrücke in der Reihenfolge ausgewertet werden, wie sie in der Funktion angegeben sind, muss der letzte logische Ausdruck alle Fälle abdecken, die von keinem anderen der vorangegangen logischen Ausdrücke akzeptiert wurden. Der einfachste logische Ausdruck, der immer wahr ist, ist der Wahrheitswert `WAHR`. Deshalb wird dieser Wert als letzter logischer Ausdruck für `WENNS()` verwendet.

```
= WENNS(A2:A3 = 1; "Eins"; A2:A3 = 3; "Drei"; A2:A3 =  4; "Vier"; A2:A3 = 8; "Acht"; WAHR; "Ungültig")
```

Diese Formel prüft die Werte in `A2:A3` auf Gleichheit mit den Werten `1`, `3`, `4` und `8`. Für diese Zahlen wird die zugehörige Zahlwert als Zeichenkette ausgegeben. Falls keiner dieser Werte gefunden wird, wird der Wert `Ungültig` zurückgegeben. 


### ERSTERWERT

Die Funktion `ERSTERWERT()` bildet einen Spezialfall von `WENNS()` ab: Wenn immer der gleiche Bereich für auf Gleichheit mit unterschiedlichen Werten überprüft werden soll, können die logischen Ausdrücke mit dieser Funktion stark vereinfacht werden. Das lässt sich am letzten Beispiel im Abschnitt `WENNS` veranschaulichen.  

Weil alle logischen Ausdrücke die Gleichheit über den gleichen Adressbereich prüfen, kann die Operation mit der Funktion `ERSTERWERT()` vereinfacht wie folgt werden.

```
= ERSTERWERT(A2:A3; 1; "Eins"; 3; "Drei"; 4; "Vier"; 8; "Acht"; "Ungültig")
```

## Rezepte

### Fehlerwerte abfangen

Viele Excel-Funktionen geben einen Fehlerwert zurück, falls die Funktion kein gültiges Ergebnis ermitteln kann. Weil sich diese Fehlerwerte in Operationen fortpflanzen, müssen diese Werte durch einen geeigneten regulären Wert ersetzt werden. Das kann mit der folgenden Entscheidung erreicht werden. 

```
= WENN( ISTFEHLER(A1); 0; A1)
```

Diese Operation ersetzt alle Fehlerwerte durch den Wert `0` und lässt alle anderen Werte unverändert.

Weil diese Entscheidung sehr oft vorkommt, gibt es die Funktion `WENNFEHLER()`, mit der die gleiche Operation einfacher ausgedrückt werden kann. 

```
= WENNFEHLER(A1; 0)
```

### Eine Zahl für genau eine Bedingung zurückgeben

Ein häufiger Spezialfall für Entscheidungen ist die Auswahl von *Zahlen*, die genau **einen** logischen Ausdrück erfüllen. Anstelle eines `WENN`-Formel lässt sich der Zielwert mit dem logischen Ausdruck multiplizieren. Der logische Ausdruck liefert `1` für `WAHR` und `0` für `FALSCH`. Die Multiplikation mit `0` liefert immer `0`. Die Multiplikation mit `1` liefert den Zielwert.

Das Beispiel gibt für die folgenden Werte alle Zahlen zurück, die grösser als 10 *und* kleiner als 20 sind.

|   A |
|:----:|
|  `13` |
|   `5` |
|  `17` |
|  `20` |
|  `12` |
|   `2` |
|  `29` |
|  `11` |
|   `7` |
|  `32` |

Normalerweise würde diese Entscheidung durch die folgende Operation abgebildet:

```         
= WENN((A1:A10 > 10) * (A1:A10 < 20); A1:A10; 0)
```

Weil alle Werte Zahlen sind, handelt es sich um einen Spezialfall. Für diesen Spezialfall lässt sich die Formel vereinfachen, indem die gesuchten Werte mit dem logischen Ausdruck multipliziert werden:

```         
= A1:A10 * (A1:A10 > 10) * (A1:A10 < 20)
```

Das Ergebnis beider Formeln sind die Werte `{13;0;17;0;12;0;0;11;0;0}`.

Damit dieses Rezept funktioniert, müssen *alle* Teile des logischen Ausdrucks genau die Werte FALSCH oder WAHR bzw. 0 oder 1 zurückgeben. Das ist notwendig, weil nur das neutrale Element die Zielwerte unverändert lässt. Eine direkte Übergabe von Zahlen im logischen Ausdruck verfälscht das Ergebnis, weil nicht mit dem *neutralen Element* gerechnet wird. 

Soll für einen logischen Ausdruck nur der Wert 1 oder 0 zurückgegeben werden, dann kann der Rückgabebereich am Anfang der Formel weggelassen werden. Es wird dann nur der logische Ausdruck angegeben. Die Formel `= (A1:A10 > 10) * (A1:A10 < 20)` hat die Werte `{1;0;1;0;1;0;0;1;0;0}` als Ergebnis.

### Fehlerwerte vergleichen

Fehlerwerte können nicht direkt mit den Vergleichsoperatoren verglichen werden, weil Excel immer den ersten gefundenen Fehlerwert als Ergebnis einer Operation zurückgibt. Deshalb müssen Fehlerwerte zuerst in normale Werte konvertiert werden. Damit verschiedene Fehlerwerte miteinander verglichen werden können, müssen die verschiedenen Fehlerwerte zuerst in eindeutige Zahlen umgewandelt werden. Das übernimmt die Funktion `FEHLER.TYP()`. Diese Zahlen können anschliessend wie gewohnt weiter verarbeitet werden.

Das folgende Beispiel weist den gegebenen Fehlerwerten eine Fehlermeldung zu:

|   A |
|:----:|
| `#NV` |
| `#WERT!` |
| `3` |

Die folgende Formel liefert die Werte `{"Fehler: #NV"; "Fehlerhafter Wert", "Kein Fehler"}`.

```
= WENNS(FEHLER.TYP(A1:A2) = 7; "Fehler: #NV"; 
        FEHLER.TYP(A1:A2) = 3; "Fehlerhafter Wert"; 
        WAHR; "Kein Fehler")
```

Weil alle Vergleiche die Gleichheit überprüfen, kann die Formel mit der Funktion `ERSTERWERT()` vereinfacht werden. Die Formel lautet dann:

```
= ERSTERWERT( WENNFEHLER(FEHLER.TYP(A1:A3); 0);
              3; "Fehlerhafter Wert";
              7; "Fehler: #NV";
              "Kein Fehler")
```

Für diesen Schritt muss die Operation mit der Funktion `WENNFEHLER()` erweitert werden, weil die Funktion `FEHLER.TYP()` einen Fehler ausgibt, wenn der übergebene Wert kein Fehlerwert ist. Weil die Fehlertypen mit Werten grösser `0` durchnummeriert sind, bietet sich für reguläre Werte der Wert `0` an. 
