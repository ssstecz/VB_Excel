Attribute VB_Name = "Módulo1"
Dim R As Integer
Dim s As Integer
Dim k As Integer
Dim x As Integer
Dim da As Date
Sub Seleciona()

Result = MsgBox("Dia mais atual?", vbYesNo + vbQuestion)
If Result = vbYes Then

k = 1
Application.ScreenUpdating = False
B3_web
Acerta_data
PosGring
Atualiza_dados
Ajusta_Graf


Else:


k = 2
Dados_B3
Application.ScreenUpdating = True
MsgBox "Copiar Dados da B3"
Application.ScreenUpdating = False
Sheets("Dados_man").Select
B3_man
Acerta_data
PosGring
Atualiza_dados
Ajusta_Graf

End If

Sheets("Pos").Select
Application.ScreenUpdating = True

End Sub
Sub B3_web()

Dim celula As String

'    Application.ScreenUpdating = False
    
   '
    'Abre site B3 paa coletar dados
    
    Sheets("Dados_web").Select
    Columns("A:F").Delete Shift:=xlToLeft
    
    '----Resolva problema do Enable Java Script
    
     With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www2.bmf.com.br/pages/portal/bmfbovespa/lumis/lum-tipo-de-participante-ptBR.asp" _
        , Destination:=Range("$A$1"))
        .Name = "ajustes-do-pregao"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    


    
    '----
    End Sub
    Sub PosGring()
    
 
    Range("A1").Select
    
    If Range("A1").Value <> "" Then
    
    Selection.QueryTable.Refresh BackgroundQuery:=False
    
    End If
    
    Cells.Find(What:="MERCADO FUTURO DE TAXA DE JURO", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    
    Cells.Find(What:="MERCADO FUTURO DE DÓLAR", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    
    R = ActiveCell.Row
    R = R + 1
    
    Do While Cells(R, 1) = ""
    R = R + 1
    Loop
    
    s = R
    d = 0
    
    Do While d <> 1
    
    If Cells(s, 1) Like "*Total*" Then d = 1
    s = s + 1
    Loop
    s = s - 1
    
    
   
    'Application.ScreenUpdating = True
    
End Sub
Sub Dados_B3()
    
Dim d As New DataObject
Dim y As String
  y = "https://www2.bmf.com.br/pages/portal/bmfbovespa/lumis/lum-tipo-de-participante-ptBR.asp"
  d.SetText y
  d.PutInClipboard
  
End Sub
Sub B3_man()
'Dim d As New DataObject
  
  
  Sheets("Dados_man").Select
  Columns("A:F").Delete Shift:=xlToLeft
 
 ' d.GetFromClipboard
  
  Range("A1").Select
  ActiveSheet.Paste
  
    
End Sub
Sub Atualiza_dados()
Attribute Atualiza_dados.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Sheets("Pos").Select
    
    If k = 1 Then
    Plan = "Dados_web"
    x = 1
    Do While Cells(x, 1) <> ""
    x = x + 1
    Loop
    Cells(x, 1).NumberFormat = "dd/mm/yyyy"
    Cells(x, 1).Value = da
   
    End If
    
    If k = 2 Then
    Plan = "Dados_man"
    x = 1
    Do While Cells(x, 1) <> da
    x = x + 1
    Loop
    
    End If
    
    
    
    
    If k <> 1 And k <> 2 Then Exit Sub
    
    
    Sheets(Plan).Cells(R, 1).Value = Sheets("Pos").Range("B1").Value
    Sheets(Plan).Cells(R + 4, 1).Value = Sheets("Pos").Range("D1").Value
    Sheets(Plan).Cells(R + 6, 1).Value = Sheets("Pos").Range("F1").Value
    'Sheets(Plan).Cells(R + 9, 1).Value = Sheets("Pos").Range("E1").Value
    
    
    
    Range("B" & x & "").FormulaR1C1 = _
        "=(VLOOKUP(R1C," & Plan & "!R" & R & "C1:R" & s & "C4,2,0))-(VLOOKUP(R1C," & Plan & "!R" & R & "C1:R" & s & "C4,4,0))"
    Range("D" & x & "").FormulaR1C1 = _
        "=(VLOOKUP(R1C," & Plan & "!R" & R & "C1:R" & s & "C4,2,0))-(VLOOKUP(R1C," & Plan & "!R" & R & "C1:R" & s & "C4,4,0))"
    Range("F" & x & "").FormulaR1C1 = _
        "=(VLOOKUP(R1C," & Plan & "!R" & R & "C1:R" & s & "C4,2,0))-(VLOOKUP(R1C," & Plan & "!R" & R & "C1:R" & s & "C4,4,0))"
    
    Range("J" & x & "").FormulaR1C1 = "=VLOOKUP(RC[-9],[CVAP_DOL_v41.xlsm]Info!C2:C14,13,0)"
    
    Range("C" & x & "").FormulaR1C1 = "=RC[-1]+R[-1]C"
    Range("E" & x & "").FormulaR1C1 = "=RC[-1]+R[-1]C"
    Range("G" & x & "").FormulaR1C1 = "=RC[-1]+R[-1]C"
    Range("H" & x & "").FormulaR1C1 = "=RC[-5]+RC[-3]"
    Range("I" & x & "").FormulaR1C1 = "=RC[-6]+RC[-4]"
    Range("K" & x & "").FormulaR1C1 = "=RC[-5]-R[-1]C[-5]"
    Range("L" & x & "").FormulaR1C1 = "=RC[-4]-R[-1]C[-4]"
    
    Range("A" & x & ":L" & x & "").Value = Range("A" & x & ":L" & x & "").Value
    
    
    Range("A" & x & ":L" & x & "").Borders(xlDiagonalDown).LineStyle = xlNone
    Range("A" & x & ":L" & x & "").Borders(xlDiagonalUp).LineStyle = xlNone
    With Range("A" & x & ":L" & x & "").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("A" & x & ":L" & x & "").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("A" & x & ":L" & x & "").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("A" & x & ":L" & x & "").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("A" & x & ":L" & x & "").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("A" & x & ":L" & x & "").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
   ' Range("B" & x & ":I" & x & "").NumberFormat = "0.00"
    Range("B" & x & ":I" & x & "").NumberFormat = "#,##0"
    Range("J" & x & "").NumberFormat = "#,##0.00"
    
   
    With Range("A" & x & ":L" & x & "")
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Range("A" & x & ":L" & x & "")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    Range("A1").Select
End Sub
Sub Acerta_data()

x = 1
Do While d <> 1
    
     If Cells(x, 1) Like "*Atualizado em*" Then d = 1
    
    x = x + 1
    Loop
    x = x - 1
    
      Cells(x, 1).TextToColumns Destination:=Range("A" & x & ""), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
        "/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)) _
        , TrailingMinusNumbers:=True
    
    If Cells(x, 4) = 1 Then Cells(x, 4) = "jan" Else
    If Cells(x, 4) = 2 Then Cells(x, 4) = "feb" Else
    If Cells(x, 4) = 3 Then Cells(x, 4) = "mar" Else
    If Cells(x, 4) = 4 Then Cells(x, 4) = "apr" Else
    If Cells(x, 4) = 5 Then Cells(x, 4) = "may" Else
    If Cells(x, 4) = 6 Then Cells(x, 4) = "jun" Else
    If Cells(x, 4) = 7 Then Cells(x, 4) = "jul" Else
    If Cells(x, 4) = 8 Then Cells(x, 4) = "aug" Else
    If Cells(x, 4) = 9 Then Cells(x, 4) = "sep" Else
    If Cells(x, 4) = 10 Then Cells(x, 4) = "oct" Else
    If Cells(x, 4) = 11 Then Cells(x, 4) = "nov" Else
    If Cells(x, 4) = 12 Then Cells(x, 4) = "dec" Else
    
    Range("F" & x & "").FormulaR1C1 = "=RC[-3]&""/""&RC[-2]&""/""&RC[-1]"
    Range("F" & x & "").NumberFormat = "dd/mm/yyyy"
    Range("F" & x & "").Value = Range("F" & x & "").Value
    da = Range("F" & x & "").Value



End Sub
Sub Ajusta_Graf()

    
    
    Sheets("Gráf1").Select

    ActiveChart.ChartArea.Select
   
    ActiveChart.FullSeriesCollection(1).Formula = "=SERIES(Pos!R1C6,Pos!R2C1:R" & x & "C1,Pos!R2C7:R" & x & "C7,1)"
    
    ActiveChart.FullSeriesCollection(2).Formula = "=SERIES(Pos!R1C8,Pos!R2C1:R" & x & "C1,Pos!R1C9:R" & x & "C9,2)"
    
    ActiveChart.FullSeriesCollection(3).Formula = "=SERIES(Pos!R1C10,Pos!R2C1:R" & x & "C1,Pos!R2C10:R" & x & "C10,3)"
    
    ActiveChart.ChartArea.Select
    
    Sheets("Gráf2").Select
    
    Sheets("Gráf2").Select
    ActiveChart.ChartArea.Select
    
    ActiveChart.FullSeriesCollection(1).Select
    Selection.Formula = "=SERIES(Pos!R1C10,Pos!R2C1:R" & x & "C1,Pos!R2C10:R" & x & "C10,1)"
    
    ActiveChart.FullSeriesCollection(2).Select
    Selection.Formula = "=SERIES(Pos!R1C11,Pos!R2C1:R" & x & "C1,Pos!R2C11:R" & x & "C11,2)"
    
    ActiveChart.FullSeriesCollection(3).Select
    Selection.Formula = "=SERIES(Pos!R1C12,Pos!R2C1:R" & x & "C1,Pos!R2C12:R" & x & "C12,3)"
    
    ActiveChart.ChartArea.Select
    
    Sheets("Gráf3").Select
    ActiveChart.ChartArea.Select
    
    ActiveChart.FullSeriesCollection(1).Select
    Selection.Formula = "=SERIES(Pos!R1C6,Pos!R2C1:R" & x & "C1,Pos!R2C7:R" & x & "C7,1)"
    
    ActiveChart.FullSeriesCollection(2).Select
    Selection.Formula = "=SERIES(Pos!R1C8,Pos!R2C1:R" & x & "C1,Pos!R2C9:R" & x & "C9,2)"
    
    ActiveChart.FullSeriesCollection(3).Select
    Selection.Formula = "=SERIES(Pos!R1C10,Pos!R2C1:R" & x & "C1,Pos!R2C10:R" & x & "C10,3)"
    
    ActiveChart.FullSeriesCollection(4).Select
    Selection.Formula = "=SERIES(Pos!R1C11,Pos!R2C1:R" & x & "C1,Pos!R2C11:R" & x & "C11,4)"
    
    ActiveChart.FullSeriesCollection(5).Select
    Selection.Formula = "=SERIES(Pos!R1C12,Pos!R2C1:R" & x & "C1,Pos!R2C12:R" & x & "C12,5)"
    
    ActiveChart.ChartArea.Select

End Sub
