# VBA-shortcutsForAudit

Makra globalne powinny być przechowywane w pliku `PERSONAL.XLSB`

Którego ściezka powinna wyglądać:
```
C:\Users\<nazwa>\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB
```
W pliku `PERSONAL.XLSB`, w module `ThisWorkbook` należy mieć inicjacje funkcji inicjującą wszystkie bindy

```
Sub Workbook_Open()
    BindShortcuts
End Sub
```
## Warto wiedzieć
Niektóre bindy nadpisują te istniejące w excelu (moim zdaniem rzadko używane)
Niektórych funkcji excela nie opłaca się bindować, z uwagi na to że występują skróty "altowe", zachowując przy tym opcje zapisu
