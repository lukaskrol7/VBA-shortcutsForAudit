Sub TogglePurpleFont()
    Dim cell As Range
    Dim rng As Range
    Dim purpleColor As Long

    purpleColor = RGB(79, 45, 127)

    Set rng = Selection

    ' loop through each cell and toggle purple + bold
    For Each cell In rng
        With cell.Font
            ' toggle off if already purple
            If .Color = purpleColor = True Then
                .Color = vbBlack      ' reset to black
                .Bold = False         ' turn off bold
            Else
                ' toggle on
                .Color = purpleColor
                .Bold = True
            End If
        End With
    Next cell
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
Sub ToggleYellow()
    Dim cell As Range
    Dim rng As Range
    Dim answer As VbMsgBoxResult

    ' current selection
    Set rng = Selection

    ' if selection is larger than 1000 cells — show warning with option to cancel, cuz u might accidently press it
    If rng.Count > 1000 Then
        answer = MsgBox("Zaznaczono " & rng.Count & " komórek." & vbCrLf & _
                        "Czy na pewno chcesz kontynuować?", _
                        vbExclamation + vbYesNo, "Uwaga: duży zakres")

        If answer = vbNo Then Exit Sub
    End If

    ' Main loop
    For Each cell In rng
        If cell.Interior.Color = vbYellow Then
            cell.Interior.ColorIndex = xlNone    ' if it already have the yellow, delete it
        Else
            cell.Interior.Color = vbYellow       ' else: make it yellow
        End If
    Next cell
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

Private Function GetAdjustedFormat(fmt As String, delta As Integer) As String
    Dim decimals As Integer
    
    ' policz kropki i zera po przecinku
    If InStr(fmt, ".") > 0 Then
        decimals = Len(Split(fmt, ".")(1))
    Else
        decimals = 0
    End If
    
    decimals = decimals + delta
    If decimals < 0 Then decimals = 0
    
    If decimals = 0 Then
        GetAdjustedFormat = "0"
    Else
        GetAdjustedFormat = "0." & String(decimals, "0")
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

