---
# bibliography: references.bib

title: Aussagenlogik

abstract: |
    Dieses Kapitel befasst sich mit der Anwendung der Aussagenlogik in Excel. Es werden die logischen Operatoren NICHT, UND, ODER und XODER behandelt. Auf dieser Grundlage werden die wichtigsten Vergleichsoperation zum Formulieren logischer Ausdrücke vorgestellt. Das Kapitel schliesst mit der Anwendung von Fallunterscheidungen.

    Es werden die folgenden Funktionen behandelt:
    `NICHT()`, `UND()`, `ODER()`, `XODER()`, `WENN()`, `WENNS()`, `ERSTERWERT()`, `WENNFEHLER()`, `FEHLER.TYP()`, `ISTFEHLER()`, `XVERWEIS()`, `IDENTISCH()`

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

Die Funktionen `UND()`, `ODER()` und `XODER()` sind *Aggregatoren*. Das bedeutet, dass Sie alle Werte in dem angegebenen Bereichen zusammenfassen. Das ist oft nicht das gewünschte Verhalten. Deshalb muss bei der Arbeit mit Vektoren auf die Bool'sche Arithmetik zurückgegriffen werden, um logische Ausdrücke richtig auszuwerten.

Zum Beispiel sollen für die folgenden Werte *paarweise* der logische Ausdruck $a \land b$ ausgewertet werden, so dass für alle Wertepaare der richtige Wahrheitswert ermittelt wird.

|   A |   B |
|:----:|:----:|
| `WAHR` | `FALSCH` |
| `FALSCH` | `WAHR` |
| `WAHR` | `WAHR` |
| `FALSCH` | `WAHR` |
| `WAHR` | `WAHR` |
| `FALSCH` | `FALSCH` |

Die Formel `= UND(A1:A6; B1:B6)` liefert den Wert `FALSCH` zurück, weil nicht alle Werte im *gesamten Bereich* von `A1:A6` und `B1:B6` gleich `WAHR` sind. Es gibt keine Funktion und keinen eigenen logischen Operator zur paarweisen logischen Verknüpfung dieser beiden Bereiche. Deshalb werden in Excel oft logische Ausdrücke in der Boole'schen Arithmethik eingesetzt, um nur die Werte aus den gleichen Datensätzen miteinander zu vergleichen. Diese Schreibweise ist immer dann notwendig, wenn logische Ausdrücke sich auf die einzelnen Datensätze beziehen.

Die Formel `= A1:A6 * B1:B6` hat die Werte `{0;0;1;0;1;0}` zum Ergebnis. Um Wahrheitswerte zuerhalten kann noch auf die Ungleichheit mit `0` geprüft werden. Dazu wird die Formel wie folgt ergänzt: `= (A1:A6 * B1:B6) <> 0`. Das Ergebnis ist nun `{FALSCH; FALSCH; WAHR; FALSCH; WAHR; FALSCH}`. Dieser Schritt ist in der Praxis selten notwendig, weil für die meisten Operationen Zahlenwerte implizit als Wahrheitswerte behandelt werden. 

Weil Excel alle Werte ungleich `0` als `WAHR` interpretiert, können die Operationen `UND()` mit `*` und die Operation `ODER()` mit `+` direkt ersetzt werden. Hierbei ist darauf zu achten, dass das nummerische Ergebnis dieser Addition oder Multiplikation *ausschliesslich* als Wahrheitswert von Bedeutung ist. 

Die Operation `XODER()` entspricht der Ungleichheit `<>`. Dabei muss allerdings darauf geachtet werden, das dieser Vergleich als Ersatz für `XODER()` *ausschliesslich* für Wahrheitswerte bzw. `0` und `1` erlaubt ist. Für die oben gezeigten Werte ergibt die Formel `= (A1:A5 <> B1:B5)` die Werte `{WAHR; WAHR; FALSCH; WAHR; FALSCH; FALSCH}`. Das Ergebnis ist deshalb sichergestellt, weil alle Vergleichswerte `0` oder `1` sind. Werden jedoch auch andere Vergleichswerte zugelassen, dann liefert die Formel `= (A1:A6 <> B1:B6)` die Werte `{WAHR; WAHR; WAHR; WAHR; WAHR; WAHR}`. Dieses Verhalten zeigt das folgende Beispiel.

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
= (A1:A6 <> 0) <> (B1:B6 <> 0)
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

Der Trick besteht darin, dass die *Rückgabematrix* durch einen Vergleich aus dem Suchbereich erzeugt wird. Dadurch wird die Rückgabematrix mit den gleichen Wahrheitswerten für alle Werte im Suchbereich gefüllt. Die erste Formel ergibt deshalb `{WAHR; FALSCH}` und die zweite Formel `{FALSCH; WAHR}`, weil der Wert `4` im Suchbereich vorkommt und der Wert `7` nicht. Diese Werte können direkt in logischen Ausdrücken verwendet werden.

### Zeichenketten vergleichen

Die Vergleichsoperatoren zeigen die Unterschiede zweier Zeichenketten bezüglich der alphabetischen Sortierung an. Die "kleinste" Zeichenkette ist die leere Zeichenkette. Gross- und Kleinschreibung wird bei Zeichenkettenvergleichen nicht unterschieden.

Weil Excel für Vergleiche die nicht-druckbaren Zeichen mit Ausnahme des Leerzeichens und des Tabulators ignoriert, gibt der Vergleichsoperator `=` `WAHR` auch für Zeichenketten zurück, die unterschiedliche nicht-druckbare Zeichen enthalten. Das gleiche Problem entsteht beim Vergleich von unterschiedlicher Gross- und Kleinschreibung. Um auch diese Unterschiede zu erkennen, müssen wir die Funktion `IDENTISCH()` verwenden. Diese Funktion vergleicht die Zeichenketten Zeichen für Zeichen und liefert nur dann `WAHR` zurück, wenn die Zeichenketten exakt gleich sind.

BEISPIEL

## Fälle unterscheiden

Logische Ausdrücke eigenen sich besonders gut, um *Fallunterscheidungen* zu formulieren, weil ein logischer Ausdruck immer nur zwei Werte als Ergebnis haben kann. Es gibt also für jeden logischen Ausdruck immer nur zwei unterscheidbare Fälle.

Excel hat zwei zentrale Funktionen für Fallunterscheidungen: `WENN()` und `WENNS()`. Die Funktion `WENN()` ist eine *einfache* Unterscheidung, die Funktion `WENNS()` unterstützt *mehrfache* Unterscheidungen. In anderen Programmiersprachen wird in diesem Zusammenhang auch von *Verzweigungen* gesprochen. 

### WENN

Die Funktion `WENN()` ist eine *einfache* Fallunterscheidung. Einfach bedeutet hier, dass die beiden Fälles *eines* logischen Ausdrucks unterschieden werden. Entsprechend hat die Funktion `WENN()` drei Parameter:

1. Der auszuwertende logische Ausdruck.
2. Das Ergebnis falls der logische Ausdruck `WAHR` ergibt.
3. Das Ergebnis falls der logische Ausdruck `FALSCH` ergibt.

Das Ergebnis für den Fall, dass der logische Ausdruck `FALSCH` ergibt, ist optional. Fehlt dieser Parameter, dann wird der Wert `FALSCH` zurückgegeben.

Das Verhalten dieser Funktion lässt sich mit den Wahrheitswerten als logischer Ausdruck direkt überprüfen: 

``` 
= WENN(WAHR; "Guten Tag"; "Auf Wiedersehen")
```

Weil der logische Ausdruck in diesem Fall `WAHR` ist, wird der zweite Parameter als Ergebnis zurückgegeben. Die Formel gibt also den Wert `"Guten Tag"` zurück.  

Wird der logische Ausdruck auf `FALSCH` geändert, dann liefert die Formel den Wert `"Auf Wiedersehen"`.

```
= WENN(FALSCH; "Guten Tag"; "Auf Wiedersehen")
```

Lassen wir den dritten Parameter weg, dann wird der Wert `FALSCH` zurückgegeben.

```
= WENN(FALSCH; "Guten Tag")
```

Ausser der Fallunterscheidung hat `WENN()` keine weiteren Eigenschaften. Deshalb wird diese Funktion in der Praxis oft mit anderen Funktionen kombiniert. Das kann mit der Funktion `WENN()` selbst geschehen. In diesem Fall wird von verschachtelten Fallunterscheidungen gesprochen.

Zur Veranschaulichung dient das folgende Beispiel:

||   A |
|----:|:----:|
|1|  `4` |
|2| `7` |

Eine Fallunterscheidung soll prüfen, ob die Werte in `A1:A2` Werte gleich `1`, `3`, `4` oder `8` sind. Falls das der Fall ist, soll der zugehörige Zahlwert als Zeichenkette ausgegeben werden. Falls das nicht der Fall ist, soll der Wert `Ungültig` zurückgegeben werden.  Als verschachtelte `WENN()`-Funktion lässt sich diese Fallunterscheidung wie folgt formulieren:

```
= WENN(A1 = 1; "Eins"; 
       WENN(A1 = 3; "Drei"; 
            WENN(A1 = 4; "Vier"; 
                 WENN(A1 = 8; "Acht"; 
                      "Ungültig"))))
```

Eine solche verschachtelte Fallunterscheidung wird als *Entscheidungsbaum* bezeichnet.

### WENNS

Die Funktion `WENN()` ist eine *einfache* Fallunterscheidung. In vielen Excel-Arbeitsmappen existieren verschachtelte Aufrufe von `WENN()`-Funktionen. Diese Aufrufe machen die Formeln nicht nur schwer lesbar, sondern auch fehleranfällig und ineffizient. Deshalb sollten verschachtelte Fallunterscheidungen unbedingt vermieden werden. Mit der Funktion `WENNS()` lassen sich verschachtelte Fallunterscheidungen vermeiden, indem alle Fallunterscheidungen in einem einzigen Funktionsaufruf zusammengefasst werden. 

> **Merke:** Verschachtelte Fallunterscheidungen mit `WENN()` unbedingt vermeiden!

Die Funktion `WENNS()` erwartet Parameterpaare, bestehend aus einem logischen Ausdruck und dem Ergebnis, falls dieser logische Ausdruck `WAHR` ergibt. Die Funktion kann bis zu 127 Parameterpaare verarbeiten, so dass sich auch sehr komplexe Fallunterscheidungen mit dieser Funktion abbilden lassen.

(*Beispiel 1*) Das folgende Beispiel zeigt die Verwendung der Funktion `WENNS()` für die verschachtelte Fallunterscheidung aus dem Abschnitt `WENN`.

```
= WENNS(A2:A3 = 1; "Eins"; 
        A2:A3 = 3; "Drei"; 
        A2:A3 = 4; "Vier"; 
        A2:A3 = 8; "Acht")
```

Das Beispiel bildet aber noch nicht die vollständige Fallunterscheidung ab. Es fehlt noch der Fall, dass keiner der logischen Ausdrücke `WAHR` ergibt. Leider kann `WENNS()` ausschliesslich logische Ausdrücke mit ihren Ergebnissen verbinden. 

> **Merke:** `WENNS()` kann nur logische Ausdrücke mit ihren `WAHR`-Ergebnissen verbinden.

Anders als bei `WENN()` gibt es keine direkte Möglichkeit, ein Ergebnis festzulegen, falls alle logische Ausdrücke `FALSCH` ergeben. Um ein solches Verhalten zu erzeugen, wird ausgenutzt, dass die Funktion `WENNS()` immer einen wahren logischen Ausdruck mit einem Ergebnis verknüpft. Weil die logischen Ausdrücke in der Reihenfolge ausgewertet werden, wie sie in der Funktion angegeben sind, muss der letzte logische Ausdruck alle Fälle abdecken, die von keinem anderen der vorangegangen logischen Ausdrücke akzeptiert wurden. Der einfachste logische Ausdruck, der immer wahr ist, ist der Wahrheitswert `WAHR`. Deshalb wird dieser Wert als letzter logischer Ausdruck für `WENNS()` verwendet.

Mit diesem Wissen lässt sich das erste Beispiel mit `WENNS()` vervollständigen:

```
= WENNS(A2:A3 = 1; "Eins"; 
        A2:A3 = 3; "Drei"; 
        A2:A3 = 4; "Vier"; 
        A2:A3 = 8; "Acht"; 
        WAHR; "Ungültig")
```

Diese Formel prüft die Werte in `A2:A3` auf Gleichheit mit den Werten `1`, `3`, `4` und `8`. Für diese Zahlen wird die zugehörige Zahlwert als Zeichenkette ausgegeben. Falls keiner dieser Werte gefunden wird, wird der Wert `Ungültig` zurückgegeben. 

Die Fallunterscheidung mit `WENNS()` endet beim ersten logischen Ausdruck, der `WAHR` ergibt. Die Funktion prüft der Reihe nach alle angegebenen logischen Ausdrücke. Sobald einer dieser Ausdrücke `WAHR` ist, wird der zugehörige Ergebniswert ausgegeben und die Funktion wird beendet. Diese Eigenschaft begründet, dass die logischen Ausdrück nur die Fälle prüfen müssen, die von den vorangegangenen logischen Ausdrücken nicht abgedeckt wurden.

(*Beispiel 2) Das folgende komplexe Beispiel zeigt, wie die Fallunterscheidung mit `WENNS()` vereinfacht werden kann.

```
WENN(J2>=O2;
    (WENN(J2>L2;
          0;
          WENN(J2<=L2;
               WENN((J2>N2)*(J2>=O2);
                    (K2+((M2-K2)/(N2-L2))*(J2-L2));
                    WENN((J2<=N2)*(J2<O2);
                         (M2+((O2-M2)/(O2-N2))*(J2-N2))
                ))
          )
    ));
    100)
```

Diese Formel ist aus zwei Gründen übermässig komplex.

1. Die Fallunterscheidung mit `WENN()` ist verschachtelt.
2. Es existieren *redundante* Fallunterscheidungen. 

Bevor die Fallunterscheidung mit `WENNS()` vereinfacht wird, werden die redundanten Fallunterscheidungen entfernt. Das betrifft die zweite (`J2 > L2`) und die vierte Fallunterscheidung (`J2 > N2`). Im jeweiligen `FALSCH`-Fall wird der gegenteilige logische Ausdruck geprüft. Das ist in diesem Fall unnötig, weil die äussere Fallunterscheidung diesen Fall bereits abdeckt. Werden die redundanten logischen Ausdrücke und unnötige Klammern entfernt, dann ergibt sich die folgende wesentlich einfachere Formel.

```
= WENN(J2>=O2;
    WENN(J2>L2;
          0;
          WENN((J2>N2)*(J2>=O2);
              K2+(M2-K2)/(N2-L2)*(J2-L2);
              M2+(O2-M2)/(O2-N2)*(J2-N2)
          )
    );
    100)
```

Die äusserste Fallunterscheidung hat für den Fall `WAHR` eine verschachtelte `WENN()`-Funktion und im Fall `FALSCH` ein einfaches Ergebnis. Das ist für `WENNS()` unhandlich, so dass die äusserste Fallunterscheidung durch Umkehrung des logischen Ausdrucks umgestellt wird. 

```
= WENN(J2<O2;
       100; 
       WENN(J2>L2;
            0;
            WENN((J2>N2)*(J2>=O2);
                 K2+(M2-K2)/(N2-L2)*(J2-L2);
                 M2+(O2-M2)/(O2-N2)*(J2-N2)
            )
       )
  )
```

Nun lassen sich die vier unterschiedlichen Fälle gut erkennen und mit `WENNS()` abbilden. Daraus ergibt sich die folgende Formel.

```
= WENNS(J2<O2; 100; 
        J2>L2; 0; 
        (J2>N2)*(J2>=O2); K2+(M2-K2)/(N2-L2)*(J2-L2); 
        WAHR; M2+(O2-M2)/(O2-N2)*(J2-N2)
  )
```

Diese Formel ist wesentlich einfacher zu lesen und zu verstehen. Beim Durchgehen der Fälle fällt auf, dass ein Teilausdruck des dritten Falls das Gegenteil des ersten Falls ist. Diese Bedingung wurde bereits im ersten Fall geprüft und würde sie nicht gelten, dann wäre die Formel bereits beendet worden. Deshalb können bereits geprüfte Teilausdrücke in den nachfolgenden Ausdrücken weggefallen. Dadurch wird nicht nur die Verschachtelung, sondern auch die Komplexität der logischen Ausdrücke vereinfacht.

```
= WENNS(J2<O2; 100; 
        J2>L2; 0; 
        J2>N2; K2+(M2-K2)/(N2-L2)*(J2-L2); 
        WAHR;  M2+(O2-M2)/(O2-N2)*(J2-N2)
  )
```

Diese Formel hat jedoch den Makel, dass der letzte Fall `WAHR` keine Konstante abbildet. Besser wäre es, wenn der zweite und der letzte Fall vertauscht wären, so dass der Wert `0` der letzte Wert ist. Dazu müssen die logischen Ausdrücke umorganisiert werden. Bei der Umorganisation ist die Reihenfolge der logischen Ausdrücke zu beachten: Die letzten beiden Fälle sind nicht umabhängig vom logischen Ausdruck `J2>L2`. Beim Umorganisieren darf diese Abhängigkeit nicht verloren gehen.

```
= WENNS(J2<O2; 100; 
        (J2<=L2)*(J2>N2); K2+(M2-K2)/(N2-L2)*(J2-L2); 
        J2<=L2; M2+(O2-M2)/(O2-N2)*(J2-N2)
        WAHR; 0
  )
```

Diese Formel ist deutlich einfacher und weniger Fehleranfällig als die ursprüngliche Formel mit verschachtelten `WENN()`-Funktionen. Es lassen sich auch weitere Fälle hinzufügen, ohne dass die Formel komplexer wird. Dabei ist zu beachten, dass diese Fälle *vor* dem Fall `WAHR` angegeben werden müssen.

### ERSTERWERT

Die Funktion `ERSTERWERT()` bildet einen Spezialfall von `WENNS()` ab: Es wird bei allen logischen Ausdrücken ein Vergleich auf Gleichheit des *Suchkriteriums* mit verschiedenen Referenzwerten durchgeführt. In diesem Fall können die logischen Ausdrücke mit `ERSTERWERT()` stark vereinfacht werden. Das lässt sich am ersten Beispiel im Abschnitt `WENNS` veranschaulichen.  

Weil alle logischen Ausdrücke die Gleichheit über den gleichen Adressbereich prüfen, kann die Operation mit der Funktion `ERSTERWERT()` vereinfacht wie folgt werden.

```
= ERSTERWERT(A2:A3; 
             1; "Eins"; 
             3; "Drei"; 
             4; "Vier"; 
             8; "Acht"; 
             "Ungültig")
```

### XVERWEIS zur Fallunterscheidung

Die Funktion `XVERWEIS()` ist als Excels Version des $\in$-Operators bereits bekannt. Die Funktion kann auch als Alternative zur Funktion `ERSTERWERT()` verwendet werden. In diesem Fall werden als Rückgabematrix keine Wahrheitswerte, sondern die Ergebnisse der Fallunterscheidung angegeben. 

Der Vorteil dieser Anwendung ist, dass die Fallunterscheidung nicht mehr auf die Anzahl der Parameterpaare beschränkt ist und die Parameterpaare zum Zeitpunkt der Formelerstellung auch nicht bekannt sein müssen. 

Das folgende Beispiel zeigt die Umsetzung des Beispiels aus dem Abschnitt `ERSTERWERT` mit `XVERWEIS()`. Dazu wird zuerst eine Tabelle mit den Vergleichswerten und den zugehörigen Ergebnissen erstellt.

|   C |   D |
|:----:|:----:|
| `1` | `Eins` |
| `3` | `Drei` |
| `4` | `Vier` |
| `8` | `Acht` |

Diese Referenztabelle stell in Spalte `C` die sog. Suchmatrix und in Spalte `D` die sog. Rückgabematrix bereit. Mit diesen Werten lässt sich die Fallunterscheidung wie folgt abbilden.

```
= XVERWEIS(A2:A3; C2:D5; D2:D5; "Ungültig")
```

Ein weiterer Vorteil von `XVERGLEICH()` gegenüber `ERSTERWERT()` sind Vergleiche mit den Operatoren `=`, `<=` oder `>=`. Diese Vergleiche lassen sich über den fünften Parameter von `XVERWEIS()` konfigurieren. Dabei steht der Wert `0` für die Gleichheit, der Wert `-1` für kleiner oder gleich und der Wert `1` für grösser oder gleich. Die Vergleiche sind immer so organisiert, dass der linke Operand dem Suchkriterium entspricht und der rechte Operand dem Wert in der Suchmatrix. Bei einem Treffer wird der Wert aus der Rückgabematrix zurückgegeben. Gibt es keinen Treffer für den Vergleich wird der Wert aus dem vierten Parameter `wenn_nicht_gefunden` geliefert.

### Anwendungshilfe für Fallunterscheidungen

Die Anwendung der verschiedenen Fallunterscheidungsfunktionen hängt von verschiedenen Kriterien ab. Diese sind hier zusammengefasst: 

Die Funktion `WENN()` wird immer dann eingesetzt, wenn ein logischer Ausdruck geprüft werden muss und nur die beiden Fälle dieses Ausdrucks unterschieden werden müssen.

Die Funktion `WENNS()` wird immer dann eingesetzt werden, wenn mehrere logische Ausdrücke geprüft werden müssen. Die logischen Ausdrücke können dabei beliebig komplex sein und sich auf verschiedene Daten und Bereiche beziehen.

Die Funktion `ERSTERWERT()` wird immer dann eingesetzt, wenn die Gleichheit des Suchkriteriums mit wenigen Referenzwerten überprüft werden soll. Die Suchkriterien sind für alle Vergleiche identisch. 

Die Funktion `XVERWEIS()` wird immer dann eingesetzt, wenn ein Vergleich auf Gleichheit, Kleiner-oder-Gleich oder Grösser-oder-Gleich durchgeführt werden muss. Die Suchkriterien und die Vergleichsoperatoren sind für alle Vergleiche identisch.

Die Funktion `XVERWEIS()` muss anstatt von `ERSTERWERT()` verwendet werden, wenn die Referenzwerte des Vergleichs zum Zeitpunkt der Formelerstellung noch nicht bekannt sind oder leicht änderbar bleiben sollen.

## Rezepte

### Fehlerwerte abfangen

Viele Excel-Funktionen geben einen Fehlerwert zurück, falls die Funktion kein gültiges Ergebnis ermitteln kann. Weil sich diese Fehlerwerte in Operationen fortpflanzen, müssen diese Werte durch einen geeigneten regulären Wert ersetzt werden. Das kann mit der folgenden Entscheidung erreicht werden. 

```
= WENN(ISTFEHLER(A1); 0; A1)
```

Diese Operation ersetzt alle Fehlerwerte durch den Wert `0` und lässt alle anderen Werte unverändert.

Weil diese Entscheidung sehr oft vorkommt, gibt es die Funktion `WENNFEHLER()`, mit der die gleiche Operation einfacher ausgedrückt werden kann. 

```
= WENNFEHLER(A1; 0)
```

### Eine Zahl für genau eine Bedingung zurückgeben

Ein häufiger Spezialfall für Unterscheidungen ist die Auswahl von *Zahlen*, die genau **einen** logischen Ausdrück erfüllen. Solche Unterscheidungen geben im `FALSCH`-Fall `0` und im anderen Fall die gesuchte Zahl zurück. In diesem  Spezialfall kann der Zielwert ohne Umweg über die `WENN()`-Funktion mit dem logischen Ausdruck multipliziert werden. 
Der logische Ausdruck liefert `1` für `WAHR` und `0` für `FALSCH`. Die Multiplikation mit `0` liefert immer `0`. Die Multiplikation mit `1` liefert den Zielwert.

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

> **Achtung:** Diese spezielle Fallunterscheidung sollte auf Korrektheit überprüft werden, wenn im `WAHR`-Fall der Wert `0` erlaubt ist. In diesem Fall wird der Wert `0` nicht vom logischen Ausdruck unterschieden.


Weil alle Werte Zahlen sind, handelt es sich um den Spezialfall, dass der `WAHR`-Wert eine Zahl und der `FALSCH`-Fall eine 0 ist. Für diesen Fall lässt sich die Formel vereinfachen, indem die gesuchten Werte mit dem logischen Ausdruck multipliziert werden:

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
