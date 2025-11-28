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

Sub FormatSheetArial()
    With ActiveSheet
        .Cells.Font.Name = "Arial"
        .Cells.Font.Size = 10
        
        .Parent.Windows(1).Zoom = 90
    End With
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

Sub AlignLeft()
    Selection.HorizontalAlignment = xlLeft
End Sub

Sub AlignRight()
    Selection.HorizontalAlignment = xlRight
End Sub

Sub AlignCenter()
    Selection.HorizontalAlignment = xlCenter
End Sub

'run below to bind shortcuts
Sub BindShortcuts()
    Application.OnKey "%{LEFT}", "AlignLeft" 'alt and aarows
    Application.OnKey "%{RIGHT}", "AlignRight"
    Application.OnKey "%{UP}", "AlignCenter"
    Application.OnKey "^+a", "FormatSheetArial"   ' Ctrl+Shift+A
    Application.OnKey "^+q", "ToggleYellow"   ' ctrl+Shift+Q
    Application.OnKey "^+i", "TogglePurpleFont"   ' Ctrl+Shift+I
    Application.OnKey "^+o", "ToggleGreenFont"   ' Ctrl+Shift+O
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

