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
Sub ToggleGreenFont()
    Dim cell As Range
    Dim rng As Range
    Dim greenColor As Long

    greenColor = RGB(0, 176, 80)   ' klasyczny Excelowy zielony

    Set rng = Selection

    ' loop through each cell and toggle green + bold
    For Each cell In rng
        With cell.Font
            ' toggle off if already green + bold
            If .Color = greenColor = True Then
                .Color = vbBlack      ' reset to balck
                .Bold = False         ' turn off bold
            Else
                ' toggle on
                .Color = greenColor
                .Bold = True
            End If
        End With
    Next cell
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

Sub BindShortcuts()

    Application.OnKey "%{LEFT}", "AlignLeft" ' Alt+Left
    Application.OnKey "%{RIGHT}", "AlignRight" ' Alt+Right
    Application.OnKey "%{UP}", "AlignCenter" ' Alt+Up
    
    Application.OnKey "^+a", "NormalizeView" ' Ctrl+Shift+A
    Application.OnKey "^+q", "ToggleYellowRed" ' Ctrl+Shift+Q
    Application.OnKey "^+i", "TogglePurpleFont" ' Ctrl+Shift+I
    Application.OnKey "^+o", "ToggleGreenFont" ' Ctrl+Shift+O
    
    Application.OnKey "^%{RIGHT}", "IncreaseDecimal" ' Ctrl+Alt+Right
    Application.OnKey "^%{LEFT}", "DecreaseDecimal" ' Ctrl+Alt+Left
    
    Application.OnKey "^+f", "SelectVisibleBlanks" ' Ctrl+Shift+F
    Application.OnKey "^+c", "ToggleCenterAcrossSelection" ' Ctrl+Shift+C

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

End Sub

