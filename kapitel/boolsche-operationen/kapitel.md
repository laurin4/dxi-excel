---
bibliography: references.bib

title: Bool'sche Operationen

abstract: ""

execute: 
  echo: false
---

Die Bool'schen Operationen sind die Grundlage für die Logik in der Programmierung. Es sind Operationen zur Verknüpfung der Wahrheitswerte `WAHR` und `FALSCH`. Die vier Grundoperationen `NICHT`, `UND`, `ODER` und `XODER` sind auch in Excel verfügbar. Die Operationen `UND`, `ODER` und `XODER` sind *Aggregatoren* und können nicht zur paarweisen Verknüpfung von Vektoren genutzt werden. Deshalb müssen wir für die Arbeit mit Vektoren auf die Bool'sche Arithmetik zurückgreifen. Dabei übersetzen wir die Wahrheitswerte als Zahlen, wobei `0` und die *leere Zelle* für `FALSCH` und alle Werte ungleich `0` sowie **alle** Zeichenketten für `WAHR` stehen.


## WENN und WENNS


