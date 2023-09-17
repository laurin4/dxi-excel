# Excel Versionen

Excel ist die Tabellenkalkulation der Microsoft Office Umgebung. Die erste Version wurde 1985 veröffentlicht und wird seitdem ständig weiterentwickelt. Aktuell existieren mehrere Versionen von Excel nebeneinander. 

- Excel365
- Excel für Desktop-Computer ohne Microsoft365
- Excel als Web-Anwendung
- Excel für Mobilgeräte

Die verschiedenen Versionen unterscheiden sich in der Funktionalität. 

Die aktuelle Version ist **Excel365** und muss auf einem Desktop-Computer oder Laptop installiert werden. Diese Version ist Teil des Microsoft365 Abonnements. Diese Version wird kontinuierlich weiterentwickelt und erhält regelmässig neue Funktionen, welche die Arbeit erleichtern. 

Teil des Microsoft365 Abonnements sind die Web- und die Mobilversionen von Excel. Beide Versionen sind in der Funktion leicht eingeschränkte Versionen von Excel365. Die Mobilversion ist für die Arbeit auf einem Smartphone oder Tablet optimiert. Die Web-Version benötigt einen Web-Browser und kann auf jedem Gerät mit Internetzugang verwendet werden. Der Unterschied zwischen Excel365 und den wird immer kleiner und die meisten Konzepte in diese Buch lassen sich auch in Excel im Web umsetzen.

Excel's Desktop Version ohne Abonnement wurde auf dem Stand von Juni 2019 eingefrohren. Diese Version wird nur noch mit Sicherheitsupdates versorgt. Neue Funktionen sind in dieser Version nicht verfügbar. Diese Version ist noch recht weit verbreitet und lässt sich nur schwer von der aktuellen Version unterscheiden, denn tatsächlich handelt es sich bei Excel365 und der Desktop-Version ohne Abonement um die selbe Anwendung. Fehlt das Abonnement, dann sind die neuen Funktionen nicht verfügbar und werden nicht angezeigt oder ausgeführt. Falls diese Version aktiv ist und eine Arbeitsmappe, die mit einer Version mit Abonement erstellt wurde, öffnet, dann erfolgt eine Fehlermeldung und die entsprechenden Funktionen werden blockiert. 

Dieses Buch setzt die aktuelle Version von Excel365 voraus. In den folgenden Kapiteln ist also immer Excel365 gemeint, wenn schlicht von Excel geschrieben wird. Die meisten Beispiele und Konzepte lassen sich in der Desktop-Version *ohne Abonnement* nicht umsetzen. Viele Beispiele lassen sich mit der Web- oder Mobilversion nachvollziehen, diesen Versionen fehlt allerdings der Importer für Daten aus anderen Dateien, der in @sec-chapter-daten-importieren behandelt wird und für viele Beispiele Voraussetzung ist.

## Excel-Versionen und Betriebssysteme

Excel365 unterscheidet sich für MacOS und Windows. Im Rahmen dieses Buchs stellt das keine Einschränkung dar, denn die Unterschiede sind gering.

Die auffälligsten Unterschiede betreffen den Importer und die bedingte Formatierung.

- Der Importer ist unter Windows etwas flexibeler gestaltet als unter MacOS. Dieser Unterschied betrifft den Inhalt dieses Buchs nicht.
- Die  *bedingte Formatierung* wird unter MacOS traditionell mit einen völlig andereren Dialog konfiguriert wird als unter Windows.

::: {.callout-warning}
## MacOS vs. Windows
In den (wenigen) Fällen, in denen sich die Excel-Versionen auf den beiden Betriebssystemen unterscheiden, wird der Unterschied durch eine *MacOS vs. Windows*-Box wie diese gesondert hervorgehoben. Das gleiche gilt für die Eigenheiten, die nur auf einem der beiden Betriebssyeme vorkommen. 
:::

::: {.callout-warning}
## MacOS

Soll unter MacOS Excel365 eine alte Version ohne Abonement ersetzen, dann muss die alte Version zuerst deinstalliert werden. Sind zusätzlich andere Office-Programme von Microsoft ohne Abonement installiert, dann müssen diese ebenfalls deinstalliert werden. Beim Deinstallieren werden die Konfigurationsdateien nicht gelöscht. Diese *müssen* manuell gelöscht werden. Erst jetzt kann die neue Version von Microsoft365 installiert werden.
:::

## Excel und Sprachen

Excel365 ist in vielen Sprachen verfügbar. Die Sprache wird von Excel vom Betriebssystem übernommen. Dieses Buch bezieht sich auf die deutschsprachige Excel Version mit Schweizer Regionseinstellungen. Alle Beispiele lassen sich in allen anderen Sprachen nachvollziehen, wenn die Werte und Funktionsnamen entsprechend angepasst werden.

::: {.callout-warning}
Die Funktionsnamen unterscheiden sich stark zwischen den einzelnen Sprachen. Oft ist es unmöglich, die Funktionsnamen durch eine Übersetzung des Funktionsnamen aus einer anderen Sprache abzuleiten. 

Beispielsweise heisst der englische Funktionsname `OFFSET()` (deutsch: *Versatz*) auf Deutsch `BEREICH.VERSCHIEBEN()` (engl. *move range*). 
:::


::: {.callout-warning}
## MacOS vs. Windows

Unter MacOS kann die Sprache von Excel365 ausschliesslich über die *Systemeinstellungen* geändert werden. Die Spracheinstellung wird vom Betriebssystem übernommen. Dadurch verändert sich die Sprache aller Programme, die auf dem Computer installiert sind. Wird die Sprache von MacOS geändert, dann muss der Computer neu gestartet werden, damit die Änderung wirksam wird. Dieser Schritt verändert jedoch nicht nur die Sprache von Menus sondern auch die Namen vieler Verzeichnisse. Deshalb wird dieser Schritt nicht empfohlen.

Unter Windows kann die Sprache nur für Excel365 geändert werden. Dazu muss in Excel unter *Datei* > *Optionen* > *Sprache* geändert werden. Die gewünschte Sprache muss unter *Office-Anwendungssprache*  als *bevorzugt* markiert werden. Die Änderung wird erst wirksam, nachdem Excel beendet und neu gestartet wurde.
:::

Die Spracheinstellung haben keinen Einfluss auf die Funktion von Excel. Die Spracheinstellung bestimmt nur die Sprache der Benutzeroberfläche. Zur Benutzeroberfläche gehören die Menüs, die Dialoge, die Hilfetexte und die **Funktionsnamen**. Wird eine Arbeitsmappe in einer Excel mit einer anderen Spracheinstellung geöffnet, dann werden die Funktionsnamen automatisch übersetzt. Es ist jedoch unmöglich die Funktionsnamen in der ursprünglichen Sprache der Arbeitsmappe einzugeben.

