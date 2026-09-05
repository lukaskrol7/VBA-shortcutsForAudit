Sub TogglePurpleFont()
    Dim rng As Range
    Dim purpleColor As Long

    purpleColor = RGB(79, 45, 127)

    ' current selection
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    ' toggle based on current state of selection
    ' (Excel checks first cell of the range)
    With rng.Font
        ' toggle off if already purple + bold
        If .Color = purpleColor And .Bold = True Then
            .Color = vbBlack      ' reset to black
            .Bold = False         ' turn off bold
        Else
            ' toggle on
            .Color = purpleColor
            .Bold = True
        End If
    End With
End Sub
Sub ToggleGreen()
    Dim rng As Range
    Dim c As Range
    Dim state As Long
    Dim greenColor As Long
    
    greenColor = RGB(0, 176, 80)   ' klasyczny Excelowy zielony
    
    Set rng = Selection
    Set c = rng.Cells(1, 1)
    
    ' odczyt stanu z pierwszej komórki zaznaczenia
    If c.Interior.ColorIndex = xlNone And c.Font.Color = greenColor Then
        state = 1
    ElseIf c.Interior.Color = greenColor And c.Font.Color = vbWhite Then
        state = 2
    Else
        state = 0
    End If
    
    Select Case state
        Case 0  ' domyslny -> zielona czcionka
            rng.Interior.ColorIndex = xlNone
            rng.Font.Color = greenColor
            rng.Font.Bold = True
        Case 1  ' biala czcionka na zielonym
            rng.Interior.Color = greenColor
            rng.Font.Color = vbWhite
            rng.Font.Bold = True
        Case 2  ' powrót do domyślnego
            rng.Interior.ColorIndex = xlNone
            rng.Font.ColorIndex = xlAutomatic
            rng.Font.Bold = False
    End Select
End Sub
Sub ToggleYellowRed()
    Dim rng As Range
    Dim c As Range
    Dim state As Long
    
    Set rng = Selection
    Set c = rng.Cells(1, 1)
    
    ' odczyt stanu z pierwszej komórki zaznaczenia -> to co będzie w pierwszym będzie traktowane jakby było w całym
    If c.Interior.Color = vbYellow And c.Font.Color = vbRed Then
        state = 1
    ElseIf c.Interior.Color = vbRed And c.Font.Color = vbYellow Then
        state = 2
    Else
        state = 0
    End If
    
    Select Case state
        Case 0  ' domyslny -> czerwony napis na żóltym
            rng.Interior.Color = vbYellow
            rng.Font.Color = vbRed
            rng.Font.Bold = True
        Case 1  ' zolty napis na czerwonym
            rng.Interior.Color = vbRed
            rng.Font.Color = vbYellow
            rng.Font.Bold = True
        Case 2  ' powrót do domyślnego
            rng.Interior.ColorIndex = xlNone
            rng.Font.ColorIndex = xlAutomatic
            rng.Font.Bold = False
    End Select
End Sub
Sub NormalizeView()
    
    Dim startSheet As Worksheet
    Set startSheet = ActiveSheet
    
    Application.ScreenUpdating = False

    Dim i As Worksheet
    For Each i In ActiveWorkbook.Worksheets
        i.Activate
        ActiveWindow.Zoom = 90
        i.Range("A1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1 'bez tych 2 linijek nie wróci widokiem do A1
    Next i

    startSheet.Activate
    Application.ScreenUpdating = True
End Sub
Sub AlignLeft()
    Selection.HorizontalAlignment = xlLeft
End Sub
Sub AlignRight()
    Selection.HorizontalAlignment = xlRight
End Sub
Sub AlignCenter()
    Selection.HorizontalAlignment = xlCenter
End Sub
Sub IncreaseDecimal()
    Dim cell As Range
    
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            cell.NumberFormat = GetAdjustedFormat(cell.NumberFormat, 1)
        End If
    Next cell
End Sub
Sub DecreaseDecimal()
    Dim cell As Range
    
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            cell.NumberFormat = GetAdjustedFormat(cell.NumberFormat, -1)
        End If
    Next cell
End Sub
    Function GetAdjustedFormat(ByVal fmt As String, ByVal delta As Long) As String 'funkcja do IncreaseDecimal() i DecreaseDecimal()
        Dim baseFmt As String
        Dim decimals As Long
        Dim posDot As Long
        Dim hasPercent As Boolean
        
        hasPercent = InStr(fmt, "%") > 0
        
        'Usuwamy %
        If hasPercent Then
            baseFmt = Replace(fmt, "%", "")
        Else
            baseFmt = fmt
        End If
        
        posDot = InStr(baseFmt, ".")
        
        If posDot > 0 Then
            decimals = Len(baseFmt) - posDot
        Else
            decimals = 0
        End If
        
        decimals = Application.Max(0, decimals + delta)
        
        If decimals > 0 Then
            baseFmt = Left(baseFmt, IIf(posDot > 0, posDot - 1, Len(baseFmt))) _
                      & "." & String(decimals, "0")
        Else
            baseFmt = Left(baseFmt, IIf(posDot > 0, posDot - 1, Len(baseFmt)))
        End If
        
        ' przywracamy %
        If hasPercent Then
            GetAdjustedFormat = baseFmt & "%"
        Else
            GetAdjustedFormat = baseFmt
        End If
    End Function
Sub SelectVisibleBlanks()
    Dim rng As Range
    Dim blanks As Range

    'zaznaczony zakres
    Set rng = Selection

    'puste komórki w zaznaczeniu
    On Error Resume Next
    Set blanks = rng.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    If blanks Is Nothing Then
        MsgBox "Brak pustych komórek w zaznaczeniu.", vbInformation
        Exit Sub
    End If

    'wybór tylko pustych widocznych
    blanks.SpecialCells(xlCellTypeVisible).Select
End Sub
Sub ToggleCenterAcrossSelection()
'Funkcja CenterAcrossSelection jako alternatywa do mergowania.
'Popularniejsza opcja w IB
    Dim rng As Range
    Dim c As Range
    Dim allCenter As Boolean
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection
    
    ' sprawdź, czy WSZYSTKIE komórki mają Center Across Selection
    allCenter = True
    For Each c In rng
        ' pomijamy scalone komórki, żeby nie wywalało błędów
        If Not c.MergeCells Then
            If c.HorizontalAlignment <> xlCenterAcrossSelection Then
                allCenter = False
                Exit For
            End If
        End If
    Next c
    
    ' jeśli wszędzie jest Center Across ->> wyłącz (wróć do ogólnego wyrównania)
    If allCenter Then
        rng.HorizontalAlignment = xlGeneral
    Else
        ' jeśli nie › ustaw Center Across Selection
        rng.HorizontalAlignment = xlCenterAcrossSelection
    End If
End Sub

Sub FindSameColors()
'pomoże do znalezienia didaskaliów/komentarzy które są w specyficznym kolorze
'do zastanowienia, czy muszą to być tło AND czionka czy tylko jedno OR nie ma sensu
    Dim ws As Worksheet
    Dim c As Range
    Dim target As Range
    
    Dim firstHit As Range
    Dim sKey As String
    Dim go As Boolean

    Set target = ActiveCell
    sKey = target.Font.Color & "|" & target.Interior.Color 'do zastanowienia

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then 'skip ukryte arkusze
            For Each c In ws.UsedRange
                If c.Address(, , , True) = target.Address(, , , True) Then  'check, czy jest na komórce początkowej ;
                                                                            'bo pętla matka przechodzi przez komórki od A1, więc żeby się nie cofać
                    go = True
                ElseIf Not c.EntireRow.Hidden And Not c.EntireColumn.Hidden Then 'do zastanowienia czy to co ukryte też wykrywać, ale raczej nie
                    If c.Font.Color & "|" & c.Interior.Color = sKey Then
                        If go Then
                            ws.Activate
                            c.Select
                            Exit Sub
                        End If
                        If firstHit Is Nothing Then Set firstHit = c
                    End If
                End If
            Next c
        End If
    Next ws

    If Not firstHit Is Nothing Then
        firstHit.Worksheet.Activate
        firstHit.Select
    End If
End Sub


Sub BindShortcuts()

    Application.OnKey "%{LEFT}", "AlignLeft" ' Alt+Left
    Application.OnKey "%{RIGHT}", "AlignRight" ' Alt+Right
    Application.OnKey "%{UP}", "AlignCenter" ' Alt+Up
    
    Application.OnKey "^+a", "NormalizeView" ' Ctrl+Shift+A
    Application.OnKey "^+q", "ToggleYellowRed" ' Ctrl+Shift+Q
    Application.OnKey "^+i", "TogglePurpleFont" ' Ctrl+Shift+I
    Application.OnKey "^+o", "ToggleGreen" ' Ctrl+Shift+O
    
    Application.OnKey "^%{RIGHT}", "IncreaseDecimal" ' Ctrl+Alt+Right
    Application.OnKey "^%{LEFT}", "DecreaseDecimal" ' Ctrl+Alt+Left
    
    Application.OnKey "^+f", "SelectVisibleBlanks" ' Ctrl+Shift+F
    Application.OnKey "^+c", "ToggleCenterAcrossSelection" ' Ctrl+Shift+C
    Application.OnKey "^+d", "FindSameColors" ' Ctrl+Shift+D

End Sub


Sub UnbindShortcuts()

    Application.OnKey "%{LEFT}"
    Application.OnKey "%{RIGHT}"
    Application.OnKey "%{UP}"
    
    Application.OnKey "^+a"
    Application.OnKey "^+q"
    Application.OnKey "^+i"
    Application.OnKey "^+o"
    
    Application.OnKey "^%{RIGHT}"
    Application.OnKey "^%{LEFT}"
    
    Application.OnKey "^+f"
    Application.OnKey "^+c"
    Application.OnKey "^+d"

End Sub

