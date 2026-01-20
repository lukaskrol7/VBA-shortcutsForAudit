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
Public Sub ToggleYellow()
    Dim rng As Range

    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    ' decyzja na podstawie aktualnego stanu zaznaczenia
    If rng.Interior.Color = vbYellow Then
        rng.Interior.ColorIndex = xlNone
        rng.Font.Color = vbNone
    Else
        rng.Interior.Color = vbYellow
        rng.Font.Color = vbRed
        rng.Font.Bold = True
    End If
End Sub

Sub FormatSheetArial()
    With ActiveSheet
        .Cells.Font.Name = "Arial"
        .Cells.Font.Size = 10
        
        .Parent.Windows(1).Zoom = 90
    End With
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

Function GetAdjustedFormat(ByVal fmt As String, ByVal delta As Long) As String
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

'run below to bind shortcuts
Sub BindShortcuts()
    Application.OnKey "%{LEFT}", "AlignLeft" 'alt and aarows
    Application.OnKey "%{RIGHT}", "AlignRight"
    Application.OnKey "%{UP}", "AlignCenter"
    'Application.OnKey "^+a", "FormatSheetArial"   ' Ctrl+Shift+A
    Application.OnKey "^+q", "ToggleYellow"   ' ctrl+Shift+Q
    Application.OnKey "^+i", "TogglePurpleFont"   ' Ctrl+Shift+I
    Application.OnKey "^+o", "ToggleGreenFont"   ' Ctrl+Shift+O
    Application.OnKey "^%{RIGHT}", "IncreaseDecimal"
    Application.OnKey "^%{LEFT}", "DecreaseDecimal"
    Application.OnKey "^+f", "SelectVisibleBlanks"
    Application.OnKey "^+c", "ToggleCenterAcrossSelection"

End Sub

Sub UnbindShortcuts()
    Application.OnKey "%{LEFT}"
    Application.OnKey "%{RIGHT}"
    Application.OnKey "%{UP}"
    Application.OnKey "^q"
    Application.OnKey "^+a"
    Application.OnKey "^+q"
    Application.OnKey "^+i"
    Application.OnKey "^%{RIGHT}"
    Application.OnKey "^%{LEFT}"
    Application.OnKey "^+o"
End Sub

