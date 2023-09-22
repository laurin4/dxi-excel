## Übungen

### Aufgabe 1: Begründen Sie, warum die Verwendung der INDEX()-Funktion in der folgenden Formel redundant ist.

```
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

### Aufgabe 1a: Passen Sie die Formel so an, dass die INDEX()-Funktion nicht mehr redundant ist, ohne die Funktion zu entfernen.

