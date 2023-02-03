Attribute VB_Name = "Módulo1"
Sub RankingCorretoras()

'
' Macro1 Macro
'

'\Rank_Corr\Hist\TXT

   ' AT = "PETR4"
    
     If AT = "PETR4" Or AT = "VALE5" Or AT Like "WIN*" Or AT Like "WDO*" Or AT Like "IND*" Then OP = 1
  '   If AT Like "*PETRK*" Or AT Like "*PETRW*" Or AT Like "*VALEK*" Or AT Like "*VALEW*" Or AT Like "*IND*" Then OP = "KW"
  '   If AT Like "*PETRL*" Or AT Like "*PETRX*" Or AT Like "*VALEL*" Or AT Like "*VALEX*" Then OP = "LX"

'07/03/17 - alterado caminho de rede de C:\Users\g0200253\Documents\bkp\BV\Rank_Corr\Hist\TXT
'para C:\Users\g0200253\Documents\bkp\BV\Mini\OP17\txt

    
    
    Workbooks.OpenText Filename:= _
        "C:\Users\g0200253\Documents\bkp\BV\Mini\OP17\txt\RC_" & AT & ".txt", Origin _
        :=65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1)), _
        TrailingMinusNumbers:=True
        
        
        
    Columns("B:L").Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Cells(1, 4).Replace What:="AM", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    Cells(1, 4).Replace What:="PM", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    H = Cells(1, 4)
    
   ' Cells(1, 13) = H
    
    'mudou de 0.4166(10h) para 0.375 (9h)
    
    If H < 0.375 Then H = H + 0.5
    
    Cells(1, 4) = H
    
    Cells(1, 3) = "TS"
    
    'Cells(1, 14) = H
    
    Range("E1").FormulaR1C1 = "=TODAY()"
    Range("E1").Value = Range("E1").Value
        Range("E1").TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("H1").FormulaR1C1 = "=CONCATENATE(""_"",RC[-2],""-"",RC[-3],""-"",RC[-1])"
    Range("H1").Value = Range("H1").Value
    D = Range("H1").Value
    Range("E1:G1").ClearContents
    
    Cells(1, 2) = AT
    
    
   ' If ULT <> 123456 Then
    If OP = 1 Then
    
    
    'alterado caminho de rede de C:\Users\g0200253\Documents\bkp\BV\Rank_Corr\Hist\TXT
    'para C:\Users\g0200253\Documents\bkp\BV\Mini\OP17\txt
    
    
        Workbooks.OpenText Filename:= _
        "C:\Users\g0200253\Documents\bkp\BV\Mini\OP17\txt\Neg_" & AT & ".txt", Origin:=65001, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1)), _
        TrailingMinusNumbers:=True
    
    
    ULT = Range("A7").Value
    
    ActiveWorkbook.Close
    
    Range("E1").Value = ULT
    Range("E1").NumberFormat = "0.00"
    
    Range("A1").Select
    
    
    End If
    
    x = 1
    
    
    Do While Cells(x, 3) <> "Qtd.Cpa."
    
    
    x = x + 1
    
    Loop
    
    x = x - 1
    
    Range("A2:A" & x & "").EntireRow.Delete
    
    Range("A3:A4").EntireRow.Delete
    
    x = 1
    
    Do While Cells(x, 1) <> ""
   
   x = x + 1
   
   Loop
   
   
    Range("B3:B" & x - 1 & "").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("G2") = "N°.Neg.  "
    Range("H2") = "Saldo"
    
    Range("A3:A" & x - 1 & "").TextToColumns Destination:=Range("A3"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(5, 1)), TrailingMinusNumbers:=True
    
  '  Columns("B:B").EntireColumn.AutoFit
  '  Rows("3:3").Select
  '  ActiveWindow.FreezePanes = True
  '  Range("A1").Select


    Do While Cells(x, 1) <> ""
    
    x = x + 1
    
    Loop
    
    x = x - 1
    
    
    
  '  x = 3
    y = 9
    
    Range(Cells(3, y), Cells(x, y + 3)).FormulaR1C1 = "=CONCATENATE(""="",RC[-6])"
    
    Range(Cells(3, y + 4), Cells(x, y + 4)).FormulaR1C1 = "=CONCATENATE(""="",RC[-6],""I"")"
    
    Range(Cells(3, y + 5), Cells(x, y + 5)).FormulaR1C1 = "=CONCATENATE(""="",RC[-6])"
    
    
   
    
    Range(Cells(3, y), Cells(x, y + 5)).Copy
    
   ' Range(Cells(3, y), Cells(x, y + 5)).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    'Range(Cells(3, y), Cells(x, y + 5)).Copy
   
    Range(Cells(3, y), Cells(x, y + 5)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    Range(Cells(3, y), Cells(x, y + 5)).Replace What:=".", Replacement:=",", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Range(Cells(3, y), Cells(x, y + 5)).Replace What:="M", Replacement:="*1000000", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range(Cells(3, y), Cells(x, y + 5)).Replace What:="B", Replacement:="*1000000000", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range(Cells(3, y), Cells(x, y + 5)).Replace What:="K", Replacement:="*1000", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
     Range(Cells(3, y), Cells(x, y + 5)).Replace What:="0K", Replacement:="*1000", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    Range(Cells(3, y), Cells(x, y + 5)).Replace What:="I", Replacement:="*1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range(Cells(3, y), Cells(x, y + 5)).Replace What:="""", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Range(Cells(3, y - 6), Cells(x, y - 1)).Value = Range(Cells(3, y), Cells(x, y + 5)).Value
    
     Range(Cells(3, y), Cells(x, y + 5)).ClearContents
    
    Range(Cells(3, y - 6), Cells(x, y - 1)).NumberFormat = "#,##0"
    
   ' Range(Cells(3, y - 6), Cells(x, y - 1)).Value = Range(Cells(3, y - 6), Cells(x, y - 1)).Value
    
    Range(Cells(3, y), Cells(x, y)).FormulaR1C1 = "=RC[-6]-RC[-4]"
    Range(Cells(3, y), Cells(x, y)).NumberFormat = "#,##0"
    
    Range(Cells(3, y + 1), Cells(x, y + 1)).FormulaR1C1 = "=IF(RC[-1]<>0,-RC[-2]/RC[-1],0)"
    Range(Cells(3, y + 1), Cells(x, y + 1)).NumberFormat = "0.00"
    
    Cells(2, 9) = "Saldo Qtd."
    Cells(2, 10) = "PM"
    
    Columns("C:J").EntireColumn.AutoFit
    
    Range("A1").Select
    
        If ULT <> 123456 Then
        
'alterado caminho de rede de C:\Users\g0200253\Documents\bkp\BV\Rank_Corr\Hist
'para C:\Users\g0200253\Documents\bkp\BV\Mini\OP17\txt\hist
        
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\g0200253\Documents\bkp\BV\Mini\OP17\txt\hist\RC_" & AT & "_" & D & "_" & H & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Application.DisplayAlerts = False
        
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\g0200253\Documents\bkp\BV\Mini\OP17\txt\hist\RC_" & AT & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
    End If
    
    If ULT = 123456 Then
    
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\g0200253\Documents\bkp\BV\Mini\OP17\txt\hist\" & OP & "\RC_" & AT & "_" & D & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
    End If
    
    Application.DisplayAlerts = True

End Sub

