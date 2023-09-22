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

::: {.callout-note}
## Merke
Tabellenadressierungen auf eine Spalte oder einen Spaltenberech sind immer Vektorisiert.
:::

> ::: {#exm-tabellen-zu-vektoren}
> ## Vektorisieren von Tabellenspalten
> 
> ```
> = Tabelle1[Spalte1]
> ```
> 
> Diese Formel vektorisiert die Spalte mit dem Namen `Spalte1` aus der Tabelle `Tabelle1`. Angenommen, dass diese Formel in Zelle `A1` des aktuellen Arbeitsblattes steht, dann kann anschliessend auf den Vektor über die Gatter-Notation zugegriffen werden: 
> 
> ```
> =A1#
> ```
> ::: 

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

Excel kann nur durch Funktionen programmiert werden. 

::: {.callout-note}
## Makros
Neben den Funktionen und Kommandos existieren in Excel noch Makros. Mit Makros können neue Kommandos und Funktionen programmiert werden. Makros folgen aber nicht den Regeln von Formeln, weil sie in einer anderen Programmiersprache geschrieben werden.

Makros unterliegen nicht den Einschränkungen von Excel-Funktionen. Diese Freiheit ist gleichzeitig ein Fluch, denn Makros sind ein Sicherheitsrisiko und Excel präsentiert entsprechende Warnungen, wenn Makros in einer Arbeitsmappe gefunden wurde. 

Aktuelle Bestrebungen von Microsoft zielen darauf ab, Makros langfristig durch Funktionen zu ersetzen. Ein Teil dieser Bestrebungen ist die Einführung von LAMBDA-Funktionen (s. [Abschnitt @sec-lambda-funktionen]).
:::

::: {.callout-note}
## Merke
Excel hat keine Identitätsfunktion. Die Identitätsfunktion wird durch eine Formel simuliert, die nur eine Adresse enthält. Solche Formeln werden für das *Vektorisieren* (@def-vektorisieren) von Bereichen eingesetzt.
:::
## Operatoren

Excel hat eine Anzahl vordefinierter Operatoren. Nicht alle Excel Operatoren haben keine direkte Entsprechung als Funktion. 

## Substitution

### Substitution über Funktionspfade

### Substitution mit `LET()`

Excels `LET()`-Funktion erlaubt das Vereinfachen komplizierter Formeln durch *Variablen*. Diese *Variablen* existieren nur im Kontext der `LET()`-Funktion und können nicht ausserhalb dieser Funktion verwendet werden.

::: {.callout-note}
## Problem
Eine Funktion wird in einer Operation mit den gleichen Parametern mehrfach aufgerufen. 
:::

Eine Variable in der LET()-Funktion entspricht einer Substitution eines Teilausdrucks einer Formel.

::: {#exm-let-substitution}
## `LET()`-Funktion zur Substitution
```Excel
=LET(
    Daten; 'Unbearbeitete Daten'!A:F;
    DatenFeld; INDEX(
        Daten;
        sequenz(ZEILEN(Daten));
        sequenz(1; SPALTEN(Daten))
      ); 
    WENN(ISTLEER(DatenFeld);#NV;DatenFeld)
 )
```
::: 

In @exm-let-substitution wird der referenzierte Bereich und der Aufruf der `INDEX()` substituiert. Die Substitution wird durch die Variablen `Daten` und `DatenFeld` realisiert. 


::: {#exm-ohne-let-substitution}
## Formel ohne Substitution

```
=WENN(ISTLEER(INDEX(
        'Unbearbeitete Daten'!A:F;
        sequenz(ZEILEN('Unbearbeitete Daten'!A:F));
        sequenz(1; SPALTEN('Unbearbeitete Daten'!A:F))
      ));
      #NV;
      INDEX(
        'Unbearbeitete Daten'!A:F;
        sequenz(ZEILEN('Unbearbeitete Daten'!A:F));
        sequenz(1; SPALTEN('Unbearbeitete Daten'!A:F))
      )
 )
 ```
:::

Die beiden Beispiele veranschaulichen, wie mit der `LET()`-Funktion mehr als eine Substitution umgesetzt wird, um eine komplexe Formel in überschaubare Teilschritte zu zerlegen und so stark zu vereinfachen.


## Funktionsketten

Normalerweise entsprechen Excel-Adressen diesem Konzept. Gerade bei Matrizen und bestimmten Transformationen ist das Zwischenspeichern von Werten auf einem Arbeitsblatt unhandlich, behindert die Übersichtlichkeit oder führt zu unerwarteten Ergebnissen. Mit der Funktion `LET()` können wir Hilfsvariablen anlegen und so Excel-Formeln stark vereinfachen. 


### LET() und leere Zellen {#sec-funktionsketten-leerezelle}

Normalerweise werden leere Zellen als Ergebnis einer Funktion durch `0` ersetzt. Dieses Konvertierung findet erst bei der Darstellung des Ergebnisses statt. Innerhalb einer Funktionskette werden leere Zellen als leere Zellen weitergereicht, solange keine Aggregation vorgenommen wird. Es ist deshalb möglich in einer Funktionskette eine Entscheidung mit `ISTLEER()` für den Fall einer leeren Zelle zu treffen.

Eine Excel-Operation **muss** einen Wert als Ergebnis einer **Formel** haben. Wird ein nicht vorhandener Wert (d.h. leere Zelle) in einem Ergebnis einer Formel gefunden, dann wird dieser Wert automatisch in den Wert `0` konvertiert. Diese Umwandlung passiert jedoch erst *nachdem* die Operation abgeschlossen ist und Excel das Ergebnis auf dem Arbeitsblatt darstellt. 

Dieses Verhalten hat zur Folge, dass solange eine Operation nicht abgeschlossen ist, die nicht vorhandenen Werte in ihrer ursprünglichen Form erhalten bleiben. Es ist also möglich diese Werte mit `ISTLEER()` zu prüfen. 

Die ursprünglichen Daten können unvollständig sein und enthalten dann leere Zellen an den entsprechenden Zellen. Diese fehlenden Werte als `0` darzustellen, kann zu verzerrten Ergebnissen führen. Deshalb sollten solche Werte mit dem *Fehler* `#NV` (lies: *Nicht Vorhanden*) markiert werden. Dieser Fehlerwert wird nicht automatisch in den Wert `0` umgewandelt, so dass die fehlenden Werte korrekt berücksichtigt werden können. 

Diese Umwandlung nutzt aus, dass Excel leere Zellen als Ergebnis von *Funktionen* zulässt, aber nicht als Ergebnis von *Formeln*. Entsprechend kann das folgende Funktionsmuster verwendet werden.


Die beiden Vektoren `G1#` und `H1#` sind Hilfsvektoren, die Sequenzen für die Zeilen- und Spaltenindizes der Datenstruktur `A:F` enthalten.

Der *logische Ausdruck* prüft, ob ein Feld mit dem Index `G1#` und `H1#` im Stichprobenobjekt leer ist. Falls das Feld in den unbearbeiteten Daten leer ist, dann wird der Wert `#NV` als Ergebnis zurückgegeben. Sonst soll der Wert im Feld übergeben. 

In dieser Operation wird die Funktion `INDEX()` zwei Mal mit den gleichen Parametern aufgerufen. Das ist unpraktisch, weil die Operation an zwei Stellen angepasst müsste, wenn die Daten mehr oder weniger Spalten haben. Besser wäre es, wenn das Zwischenergebnis der `INDEX()`-Funktion aus der Operation herausgelöst wird und über eine Funktionsverkettung eingebunden wird. Das ist aber nicht möglich, weil Excel bei diesem Zwischenschritt die fehlenden Werte in `0` ändert, sodass anschliessend der logische Ausdruck immer `FALSCH` liefern würde.

Mittels der `LET()` Funktion wird das Ergebnis dieses Zwischenschritts in einer temporären *Variablen* gespeichert. Gegenüber der normalen Funktionsverkettung durch Funktionspfade hat diese Strategie den Vorteil, dass für Excel die Operation nicht abgeschlossen ist und deshalb die fehlenden Werte *noch nicht* in den Wert `0` umgewandelt werden. Der logische Ausdruck in der `WENN()`-Funktion kann also `WAHR` ergeben, wenn in den Daten ein Wert fehlt. Ausserdem muss die Indizierung für eine Position nur einmal durchgeführt werden, was bei komplexen Formeln die Übersichtlichkeit erhöht und die Ausführung beschleunigt.

Die ursprüngliche Formel lässt sich also dahingehend vereinfachen, dass der Aufruf der `INDEX()`-Funktion "ausgeklammert" und in der Hilfsvariablen `Feld` gespeichert wird. 

Daraus ergibt sich die Lösung als Funktionskette.

```
=LET(Feld; INDEX('Unbearbeitete Daten'!A:F;A2#;B1#); 
     WENN(
        ISTLEER(Feld);
        #NV;
        Feld
     )
 )
```

Diese Lösung entspricht ungefähr der Funktionskette @eq-funktionskette-let.

$$
Index() \triangleright Wenn()
$$ {#eq-funktionskette-let}

Damit wird der Aufruf der `WENN()`-Funktion vereinfacht, weil nur noch die Hilfsvariable `Feld` übergeben müssen. Diese Variable enthält die Daten für die Funktionsverkettung, so dass eine zusätzliche Arbeitsblattadresse nicht notwendig ist. 

## Funktionen selbst definieren {#sec-lambda-funktionen}

LAMBDA Funktion

- Funktionen können Funktionsketten beinhalten. 
- Funktionen können nur Bereiche als Ergebnis erzeugen.
- Funktionen können auch intern keine komplexeren Datenstrukturen erzeugen, als Excel normalerweise zulässt. Es sind maximal 2D Bereiche erlaubt. Komplexere Sturkturen müssen über Hilfsvariablen abgebildet werden.

Strategie der Funktionsentwicklung

- Funktion als Funktionspfad
- Funktion als Funktionskette
- Funktion als LAMBDA Funktion
- Funktion als benanntes Element
