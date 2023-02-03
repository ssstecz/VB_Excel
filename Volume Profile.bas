Attribute VB_Name = "Módulo1"
Sub Atualizar()

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd

    .Filters.Clear
    .AllowMultiSelect = False
    .InitialFileName = "C:\Users\ssste\Documents\Trd\TT23\"

    If .Show = True Then

        File = .SelectedItems(1)

    End If
    
      If File = "" Then Exit Sub

End With


Workbooks.Open Filename:= _
        "" & File & ""
    
    If Range("A1").Value Like "*Neg*" Then
    
    Orig_Tryd
    Else: Orig_B3
    
    End If

    

End Sub
Sub Orig_Tryd()

Application.ReferenceStyle = xlA1
    
    Range("A1").EntireRow.Delete
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1)), TrailingMinusNumbers:=True
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    nt = 1
    Do While Cells(nt, 1) <> ""
    nt = nt + 1
    Loop
    nt = nt - 1
    
    Range("A1").Select
    Selection.EntireColumn.Insert
    
    Range("A1").FormulaR1C1 = "DIA"
    Range("H1").FormulaR1C1 = "VOL"
    
    '----V2 DATA
    
     Range("X1") = ActiveSheet.Name
    
    Range("X1").TextToColumns Destination:=Range("X1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
        "-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)) _
        , TrailingMinusNumbers:=True
    Range("AD1").FormulaR1C1 = "=RC[-3]&""/""&RC[-2]&""/""&RC[-1]"
    Range("AD1").NumberFormat = "dd/mm/yyyy"
    
   Range("AE1").FormulaR1C1 = "=RC[-1]+0"
    
    da = Range("AE1").Value
    Range("W1").FormulaR1C1 = "=UPPER(RC[3])"
    serie = Range("W1").Value
    
    
    '----
    
    Range("A2:A" & nt & "").NumberFormat = "dd/mm/yyyy"
    
    Range("A2:A" & nt & "") = da
    
        
    Columns("A:A").EntireColumn.AutoFit
    
    Range("A2:A" & nt & "").Value = Range("A2:A" & nt & "").Value

    
    Range("H2:H" & nt & "").FormulaR1C1 = _
        "=IF(RC[-1]=""C"",RC[-5],(IF(RC[-1]=""V"",-RC[-5],0)))"
    

    Range("I2:J" & nt & "").FormulaR1C1 = _
        "=(RC[-6]+0)"
     Range("C2:D" & nt & "").Value = Range("I2:J" & nt & "").Value
    Range("D2:D" & nt & "").NumberFormat = "#,##0.00"
    Range("I2:J" & nt & "").Value = ""
 
    Range("AA1").Value = "T"
    
    VAPv3
 

End Sub
Sub Orig_B3()

 Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 9), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 9), Array(9, 9)), TrailingMinusNumbers:=True
        
        nt = 2
    Do While Cells(nt, 1) <> ""
    nt = nt + 1
    Loop
    nt = nt - 1
    
    Range("X1").Value = Range("B2").Value
    Range("Y1").FormulaR1C1 = "=LOWER(LEFT(RC[-1],3))"
    ativo = Range("Y1").Value
    Range("Z1").FormulaR1C1 = "=RIGHT(RC[-2],3)"
    serie = Range("Z1").Value
    
    Range("I2:I" & nt & "").FormulaR1C1 = _
        "=(IF(LEFT(RC[-4],1)=""9"",LEFT(RC[-4],1),LEFT(RC[-4],2))&"":""&(IF(LEFT(RC[-4],1)=""9"",MID(RC[-4],2,2),MID(RC[-4],3,2)))&"":""&(IF(LEFT(RC[-4],1)=""9"",MID(RC[-4],4,2),MID(RC[-4],5,2))))"
    
    
    Range("B2:B" & nt & "").Value = Range("I2:I" & nt & "").Value
    Range("E2:E" & nt & "").Value = ""
    Range("I2:I" & nt & "").Value = ""
    
    Range("C1").EntireColumn.Insert
    
    Range("A1").Value = "DIA"
    Range("B1").Value = "Hora"
    Range("C1").Value = "Qtd"
    Range("D1").Value = "Preço"
    Range("E1").Value = "CC"
    Range("F1").Value = "CV"
    Range("G1").Value = "Agr"
    Range("H1").Value = "VOL"
    Range("I1").Value = "NEG"
    
    Columns("A:A").EntireColumn.AutoFit
    
    Range("I2:I" & nt & "").FormulaR1C1 = "=RC[-2]/10"
    Range("I2:I" & nt & "").Value = Range("I2:I" & nt & "").Value
    Range("G2:G" & nt & "").Value = ""
    
    Range("C2:C" & nt & "").NumberFormat = "#,##0"
    Range("C2:C" & nt & "").Value = Range("E2:E" & nt & "").Value
    Range("E2:E" & nt & "").Value = ""
    
    
    Range("X1").EntireColumn.Delete
    Range("AA1").Value = "B"
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    Range("A1").Select
    
    VAPv3

End Sub
Sub VAPv3()
    
    plat = ActiveWorkbook.Name
    pl = ActiveSheet.Name
    ativo = Range("Y1").Value
    serie = Range("Z1").Value
    da = Range("A2").Value
    
    nt = 2
    Do While Cells(nt, 1) <> ""
    nt = nt + 1
    Loop
    nt = nt - 1
    
    
    Range("A1:H" & nt & "").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "" & pl & "!R1C1:R" & nt & "C8", Version:=xlPivotTableVersion15). _
        CreatePivotTable TableDestination:="Plan1!R3C1", TableName:= _
        "Tabela dinâmica1", DefaultVersion:=xlPivotTableVersion15
    Sheets("Plan1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DIA")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Preço")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("Tabela dinâmica1").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica1").PivotFields("Qtd"), "Soma de Qtd", xlSum

    ActiveSheet.PivotTables("Tabela dinâmica1").PivotSelect "Preço[All]", _
        xlLabelOnly + xlFirstRow, True
    With ActiveSheet.PivotTables("Tabela dinâmica1")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    Range("A5").Select
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("DIA").RepeatLabels = _
        True
    
    l = 5
    Do While Cells(l, 1) <> ""
    l = l + 1
    Loop
    l = l - 3
'-----INCLUI SÉRIE NA COLUNA C E DATA DO DIA NA COLUNA D, REPETE DADOS DE PREÇO E VOLUME NAS COLUNAS E / F. DEPOIS COPIARA PARA PLANILHA CONSOLIDADA

Range("A5:C" & l & "").Select

ntab = 3

Windows("CVAP_" & ativo & "_v41.xlsm").Activate
Sheets("TT_Neg").Select

Do While Cells(ntab, 4) <> ""

ntab = ntab + 1

Loop

ntab = ntab - 1


ActiveSheet.ListObjects("Tabela1").Resize Range("$A$3:$G$" & ntab + l - 4 & "")

Windows(plat).Activate
Range("A5:C" & l & "").Copy

Windows("CVAP_" & ativo & "_v41.xlsm").Activate

Range("E" & ntab + 1 & "").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


Range("D" & ntab + 1 & ":d" & ntab + l - 4 & "").Value = serie

'--------CONCLUÍDA INCLUSÃO DE VOLUME X PREÇO
    
    
    Windows(plat).Activate
       
    Sheets(pl).Select
    
    
    '------- Res Neg
    
    mcc = 2
  
    Do While Cells(mcc, 2) <> ""
  
    mcc = mcc + 1
   
    Loop
    
    '------- Define Qtd. de negócios para cálculo
              
    qq = 1000
    
    
    Range("J2").Value = 1
  
    '-----------------
    
    If Range("AA1").Value = "T" Then
    
    Range("I2:I" & mcc & "").FormulaR1C1 = _
         "=IF(RC[-1]=0,(IF(RC[-2]<>""CR"",0,R[-1]C9+1)),R[-1]C9+1)"
    End If
    
   '-------------
   
    If Range("AA1").Value = "B" Then
   
   Range("G2").Value = "L"
   Range("I2:I" & mcc & "").Value = ""
   Range("I2").Value = 0
   
   Range("G3:G" & mcc & "").FormulaR1C1 = _
        "=IF(AND(R[-1]C[-3]=RC[-3],RC[-5]=R[-1]C[-5],R[-1]C[2]=0),""L"",""X"")"
   
   Range("I3:I" & mcc & "").FormulaR1C1 = "=IF(RC[-2]=""X"",R[-1]C+1,0)"
    
    End If
   
   '-------------
   
   '-----------------
   
    Range("J3:J" & mcc & "").FormulaR1C1 = _
        "=IF(MOD((R[-1]C9+1)," & qq & ")=0,(QUOTIENT((R[-1]C9+1)," & qq & ")+1),"""")"
 
   
   
    Range("I1").Value = "NEG"
    Range("J1").Value = "CL"
    Range("K1").Value = "TS"
    Range("L1").Value = "VAL"
    
    Range("K2:K" & mcc & "").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10]+RC[-9],"""")"
    Range("L2:L" & mcc & "").FormulaR1C1 = "=IF(RC[-2]<>"""",RC[-8],"""")"
             
    '------cálculo da média de extensão em qq (não qq/5 , alterado em 05/01/23)
    
    If ativo = "ind" Or ativo = "dol" Then
    
    ly = 2
    
    Do While Range("A" & ly & "") = 0
    ly = ly + 1
    Loop
    
    lx = 200
    
    lx = lx + ly
    Do While lx < mcc
   
    Range("M" & ly & ":M" & lx & "").FormulaR1C1 = "=IF(MOD((RC9)," & qq / 5 & ")=0,(LARGE(R" & ly & "C[-9]:RC[-9],1)),"""")"
    Range("N" & ly & ":N" & lx & "").FormulaR1C1 = "=IF(MOD((RC9)," & qq / 5 & ")=0,(SMALL(R" & ly & "C[-10]:RC[-10],1)),"""")"
    Range("O" & ly & ":O" & lx & "").FormulaR1C1 = "=IF(MOD((RC9)," & qq / 5 & ")=0,(ABS((RC[-2]-RC[-1]))),"""")"
     
     
     If Cells(lx - 1, 9) > 200 Then
    
    Range("P" & ly & ":P" & lx & "").FormulaR1C1 = "=IF(MOD((RC9)," & qq / 5 & ")=0,(RC[-14]-R[-199]C[-14]),"""")"
   
    End If
     
     
     lx = lx + 200
     ly = ly + 200
     
    
     Loop
     
     lz = mcc - (2 * mcc) + 2
     
     Range("O" & mcc & "").FormulaR1C1 = "=AVERAGE(R[" & lz & "]C:R[-1]C)"
     Range("O" & mcc + 1 & "").FormulaR1C1 = "=LARGE(R[" & lz - 1 & "]C:R[-2]C,1)"
     Range("O" & mcc + 2 & "").FormulaR1C1 = "=SMALL(R[" & lz - 2 & "]C:R[-3]C,1)"
     
     Range("P" & mcc & "").FormulaR1C1 = "=AVERAGE(R[" & lz & "]C:R[-1]C)"
     Range("P" & mcc + 1 & "").FormulaR1C1 = "=LARGE(R[" & lz - 1 & "]C:R[-2]C,1)"
     Range("P" & mcc + 2 & "").FormulaR1C1 = "=SMALL(R[" & lz - 2 & "]C:R[-3]C,1)"
   
     
     mr = Range("O" & mcc & "").Value
     maxrot = Range("O" & mcc + 1 & "").Value
     minrot = Range("O" & mcc + 2 & "").Value
     
     trm = Range("P" & mcc & "").Value
     trmax = Range("P" & mcc + 1 & "").Value
     trmin = Range("P" & mcc + 2 & "").Value
     
     End If
   
        
    '-----
 
  
     linp = mcc
     
     Do While Range("J" & linp & "").Value = ""
     linp = linp - 1
     Loop
     np = Range("J" & linp & "").Value
    
    
    
    Range("J" & mcc - 1 & "").Value = np + 1
            
            Range("I" & mcc & ":P" & mcc & "").Value = ""
            
             
            'valor ohlc fechamento
            
    
    Range("H1:L" & mcc - 1 & "").Value = Range("H1:L" & mcc - 1 & "").Value
        
    
    Rows("1:1").AutoFilter
    
    ActiveSheet.Range("$A$1:$L$" & mcc - 1 & "").AutoFilter Field:=11, Criteria1:="<>"
  
    
    
    Range("K2:L" & mcc - 1 & "").Copy
    
    
    
    Windows("CVAP_" & ativo & "_v41.xlsm").Activate
    
    Sheets("TT_Neg").Select
 
    


    Range("A" & ntab + 1 & "").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
     ntab1 = ntab + 1
     ntn = ntab + l - 3
    
  
     If Cells(ntn, 4) = "" And Cells(ntn, 1) <> "" Then
 
    
   
    Do While Cells(ntn, 1) <> ""

   
     ntn = ntn + 1

    Loop

 
    ntn = ntn - 1
    
    ActiveSheet.ListObjects("Tabela1").Resize Range("$A$3:$K$" & ntn & "")
    
    Range("D" & ntab + l - 4 & ":D" & ntn & "").Value = serie
    Range("E" & ntab + l - 4 & ":E" & ntn & "").Value = da
    
    Range("C" & ntab + 1 & "").Value = 0
    Range("C" & ntab + 2 & ":C" & ntn & "").FormulaR1C1 = "=RC[-2]-R[-1]C[-2]"
   
    
    Range("C" & ntab + 1 & ":C" & ntn & "").NumberFormat = "mm:ss;@"
    Range("C" & ntab + 1 & ":C" & ntn & "").Value = Range("C" & ntab + 1 & ":C" & ntn & "").Value
    
    
 
   

End If
    
    
    
    Range("C" & ntab + 1 & "").Value = 0
    Range("C" & ntab + 2 & ":C" & ntab + l - 4 & "").FormulaR1C1 = "=RC[-2]-R[-1]C[-2]"
   
    
    Range("C" & ntab + 1 & ":C" & ntab + l - 4 & "").NumberFormat = "mm:ss;@"
    Range("C" & ntab + 1 & ":C" & ntab + l - 4 & "").Value = Range("C" & ntab + 1 & ":C" & ntab + l - 4 & "").Value
    
    If ntn <= ntab + l - 3 Then
    ntemp = ntab + 1
    Do While Cells(ntemp, 2) <> ""
    ntemp = ntemp + 1
    Loop
   
   Range("C" & ntemp & ":C" & ntab + l - 4 & "").Value = ""
    
    End If
    
     
   ActiveWorkbook.RefreshAll
    
    
   tvt1 = 5
    
   Do While Sheets("Comp").Cells(tvt1, 7) <> ""
   tvt1 = tvt1 + 1
   Loop
   tvt1 = tvt1 - 1
    
   Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(tvt1, 3)).FormulaR1C1 = "=RC[-1]"
   Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(tvt1, 4)).FormulaR1C1 = "=RC[-3]"
    
    
    
   tvt = 5
    
   Do While Sheets("Comp").Cells(tvt, 7) <> ""
   tvt = tvt + 1
   Loop
   tvt = tvt - 1
    
   Range(Sheets("Comp").Cells(5, 9), Sheets("Comp").Cells(tvt, 11)).FormulaR1C1 = "=RC[-2]"
    
   
    Range("A" & ntab + 1 & "").Select
    
    
    '-----insere dados adicionais na planilha Info
    
    If ativo = "ind" Or ativo = "dol" Then
    
    
    Sheets("Info").Select
    
    inf1 = 3
    Do While Sheets("Info").Cells(inf1, 2) <> ""
    
    inf1 = inf1 + 1
    Loop
    
    Sheets("Info").Cells(inf1, 2) = da
    Sheets("Info").Cells(inf1, 2).NumberFormat = "dd/mm/yyyy"
    
    Sheets("Info").Cells(inf1, 3) = mr
    Sheets("Info").Cells(inf1, 4) = maxrot
    Sheets("Info").Cells(inf1, 5) = minrot
    
    If ativo = "dol" Then
    
    Range(Sheets("Info").Cells(inf1, 3), Sheets("Info").Cells(inf1, 5)).NumberFormat = "0.00"
    
    End If
    
    If ativo = "ind" Then
    
    Sheets("Info").Cells(inf1, 3).NumberFormat = "0.00"
    Range(Sheets("Info").Cells(inf1, 4), Sheets("Info").Cells(inf1, 5)).NumberFormat = "#,##0"
    
    End If
    
    
    Sheets("Info").Cells(inf1, 6) = trm
    Sheets("Info").Cells(inf1, 7) = trmax
    Sheets("Info").Cells(inf1, 8) = trmin
    Range(Sheets("Info").Cells(inf1, 6), Sheets("Info").Cells(inf1, 8)).NumberFormat = "mm:ss"
     
    Range(Sheets("Info").Cells(inf1, 2), Sheets("Info").Cells(inf1, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(inf1, 2), Sheets("Info").Cells(inf1, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(inf1, 2), Sheets("Info").Cells(inf1, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(inf1, 2), Sheets("Info").Cells(inf1, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(inf1, 2), Sheets("Info").Cells(inf1, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(inf1, 2), Sheets("Info").Cells(inf1, 8)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
   
    Range(Sheets("Info").Cells(inf1, 2), Sheets("Info").Cells(inf1, 8)).HorizontalAlignment = xlCenter
    Range(Sheets("Info").Cells(inf1, 2), Sheets("Info").Cells(inf1, 8)).VerticalAlignment = xlCenter

    
'    Sheets("Info").Cells(1, 1).Select
     
    End If
   
   Escala
    
    '---- COPIA VALORES PARA PLANILHA CONSOLIDADA
    
    
    
   Range("A1").Select
    
    
    Windows(plat).Activate
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
   ' Application.ScreenUpdating = True
    
End Sub
Sub Escala()

'---- acertar valores conforme o filtro
    
    '----
    
    
    tvt1 = 5
    
   Do While Sheets("Comp").Cells(tvt1, 2) <> ""
   tvt1 = tvt1 + 1
   Loop
   tvt1 = tvt1 - 1
    
   Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(tvt1, 3)).FormulaR1C1 = "=RC[-1]"
   Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(tvt1, 4)).FormulaR1C1 = "=RC[-3]"
    
    
    
   tvt = 5
    
   Do While Sheets("Comp").Cells(tvt, 7) <> ""
   tvt = tvt + 1
   Loop
   tvt = tvt - 1
    
   Range(Sheets("Comp").Cells(5, 9), Sheets("Comp").Cells(tvt, 11)).FormulaR1C1 = "=RC[-2]"
   
    
    
    
    '----
    
    gm1 = 4
    
    Do While Sheets("Comp").Cells(gm1, 1) <> ""
    gm1 = gm1 + 1
    Loop
    gm1 = gm1 - 1
    
    
    gm2 = 4
    
    Do While Sheets("Comp").Cells(gm2, 7) <> ""
    gm2 = gm2 + 1
    Loop
    gm2 = gm2 - 1
    
    
    max1 = Application.WorksheetFunction.Max(Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4)))
    min1 = Application.WorksheetFunction.Min(Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4)))
    
    
    max2 = Application.WorksheetFunction.Max(Range(Sheets("Comp").Cells(5, 10), Sheets("Comp").Cells(gm2, 10)))
    min2 = Application.WorksheetFunction.Min(Range(Sheets("Comp").Cells(5, 10), Sheets("Comp").Cells(gm2, 10)))
    
    
    If max1 > max2 Then maxgm = max1 Else maxgm = max2
    If min1 < min2 Then mingm = min1 Else mingm = min2
    
    
    Sheets("Graf").Activate
    
    Range("X12") = maxgm
    Range("X13") = mingm
    Range("X14") = maxgm - mingm
    Range("X15") = (maxgm + mingm) / 2
    Range("X16") = (Application.WorksheetFunction.SumProduct(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)), (Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4))))) / (Application.Sum(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3))))
    
    
   ' k1 = Application.WorksheetFunction.SumProduct(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)), (Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4))))
   ' k2 = Application.Sum(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)))
   ' k3 = k1 / k2
    
    
    desvpv = WorksheetFunction.StDev_P(Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4)))

    
    vt = Application.Sum(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)))
    vva = vt * 0.707
    
    Range("X18") = desvpv
    
    vp = Application.WorksheetFunction.Max(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)))
    
    lvp = 5
    Do While Sheets("Comp").Cells(lvp, 3) <> vp
    lvp = lvp + 1
    Loop

    POC = Sheets("Comp").Cells(lvp, 4)
    
    Range("X8") = POC
    
    
    lval = lvp - 1
    lvah = lvp + 1

    vcv = vp

    Do While vcv < vva

    vcv = vp + WorksheetFunction.Sum(Range(Sheets("Comp").Cells(lvp - 1, 3), Sheets("Comp").Cells(lval, 3))) + WorksheetFunction.Sum(Range(Sheets("Comp").Cells(lvp + 1, 3), Sheets("Comp").Cells(lvah, 3)))

    If Sheets("Comp").Cells(lval - 2, 3) <> "" Then
    lval = lval - 1
    End If

    If Sheets("Comp").Cells(lvah + 1, 1) <> "" Then
    lvah = lvah + 1
    End If

    Loop


    lval = lval + 1
    lvah = lvah - 1
    
    If lval = 5 Then lval = 6

    VALOW = Sheets("Comp").Cells(lval, 4)
    VAHI = Sheets("Comp").Cells(lvah, 4)
    
    
    Range("X7") = VAHI
    Range("X9") = VALOW
    
    Range("X20") = Sheets("Comp").Cells(5, 10)
    Range("X21") = Sheets("Comp").Cells(gm2, 10)
    
   tvt = 5
    
   Do While Sheets("Comp").Cells(tvt, 12) <> ""
   tvt = tvt + 1
   Loop
   tvt = tvt - 1
   
   Range(Sheets("Comp").Cells(5, 12), Sheets("Comp").Cells(tvt, 14)) = ""
 
    
    Range("Comp!$L$5:$L$" & gm2 & "") = VALOW
    Range("Comp!$M$5:$M$" & gm2 & "") = POC
    Range("Comp!$N$5:$N$" & gm2 & "") = VAHI
    
    Sheets("Graf").Select
   
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(2).Values = "=Comp!$C$5:$C$" & gm1 & ""
    ActiveChart.SeriesCollection(2).XValues = "=Comp!$D$5:$D$" & gm1 & ""
    ActiveChart.SeriesCollection(1).Values = "=Comp!$J$5:$J$" & gm2 & ""
    ActiveChart.SeriesCollection(1).XValues = "=Comp!$G$5:$G$" & gm2 & ""
    ActiveChart.SeriesCollection(3).Values = "=Comp!$N$5:$N$" & gm2 & ""
    ActiveChart.SeriesCollection(3).XValues = "=Comp!$G$5:$G$" & gm2 & ""
    ActiveChart.SeriesCollection(4).Values = "=Comp!$M$5:$M$" & gm2 & ""
    ActiveChart.SeriesCollection(4).XValues = "=Comp!$G$5:$G$" & gm2 & ""
    ActiveChart.SeriesCollection(5).Values = "=Comp!$L$5:$L$" & gm2 & ""
    ActiveChart.SeriesCollection(5).XValues = "=Comp!$G$5:$G$" & gm2 & ""
    
    
   '---- acertar valores da escala
    
    ActiveChart.Axes(xlValue).Select
    
    ActiveChart.Axes(xlValue).MaximumScale = maxgm
    ActiveChart.Axes(xlValue).MinimumScale = mingm
    

    ActiveChart.ChartArea.Select
    Range("A1").Select
    


End Sub
Sub Val_VA()
'
' Macro2 Macro
'

'
    Range("A1").Select
    
    va = 3
    Do While Cells(va, 10) <> ""
    va = va + 1
    Loop
   
    
    Range("I" & va & "").FormulaR1C1 = "=Graf!R10C24"
    Range("J" & va & "").FormulaR1C1 = "=Graf!R9C24"
    Range("K" & va & "").FormulaR1C1 = "=Graf!R8C24"
    Range("L" & va & "").FormulaR1C1 = "=Graf!R7C24"
    Range("M" & va & "").FormulaR1C1 = "=Graf!R6C24"
    Range("N" & va & "").FormulaR1C1 = "=Graf!R21C24"
    Range("O" & va & "").FormulaR1C1 = "=Graf!R14C24"
    Range("P" & va & "").FormulaR1C1 = "=Graf!R11C24"
    
    Range("I" & va & ":P" & va & "").Value = Range("I" & va & ":P" & va & "").Value
    
    
    
    If Range("I" & va & "").Value < 20000 Then
    
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).NumberFormat = "#,##0.00"
    
    End If
    
    If Range("I" & va & "").Value > 20000 Then
    
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).NumberFormat = "#,##0"
    
    End If
    
    
     
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).Borders(xlInsideVertical).LineStyle = xlContinuous
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
   
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).HorizontalAlignment = xlCenter
    Range(Sheets("Info").Cells(va, 9), Sheets("Info").Cells(va, 16)).VerticalAlignment = xlCenter
    
    Range("C1:H1").FormulaR1C1 = "=AVERAGE(R[3]C:R" & va & "C)"
    Range("O1:P1").FormulaR1C1 = "=AVERAGE(R[3]C:R" & va & "C)"
 '   Range("D1").FormulaR1C1 = "=AVERAGE(R[3]C:R" & va & "C)"
 '   Range("E1").FormulaR1C1 = "=AVERAGE(R[3]C:R" & va & "C)"
 '   Range("F1").FormulaR1C1 = "=AVERAGE(R[3]C:R" & va & "C)"
    
    '------Atualiza Gráf 2 - ULT  VA
    
   ' ActiveSheet.ChartObjects("Gráfico 2").Activate
    Sheets("Graf2").ChartObjects("Gráfico 2").Activate

    ActiveChart.FullSeriesCollection(1).Values = "=Info!$N$4:$N$" & va & ""
    ActiveChart.FullSeriesCollection(2).Values = "=Info!$M$4:$M$" & va & ""
    ActiveChart.FullSeriesCollection(3).Values = "=Info!$L$4:$L$" & va & ""
    ActiveChart.FullSeriesCollection(4).Values = "=Info!$K$4:$K$" & va & ""
    ActiveChart.FullSeriesCollection(5).Values = "=Info!$J$4:$J$" & va & ""
    ActiveChart.FullSeriesCollection(6).Values = "=Info!$I$4:$I$" & va & ""
       
    
    gg = 1
    Do While gg < 7
    ActiveChart.FullSeriesCollection(gg).XValues = "=Info!$B$4:$B$" & va & ""
    gg = gg + 1
    Loop
    
    
    ActiveChart.Axes(xlValue).MinimumScale = (Application.WorksheetFunction.Min(Range(Sheets("Info").Cells(4, 9), Sheets("Info").Cells(va, 13)))) * 0.99
    ActiveChart.Axes(xlValue).MaximumScale = (Application.WorksheetFunction.Max(Range(Sheets("Info").Cells(4, 9), Sheets("Info").Cells(va, 13)))) * 1.01
        
     uu = (Application.WorksheetFunction.Min(Range(Sheets("Info").Cells(4, 9), Sheets("Info").Cells(va, 13)))) * 0.99
    vv = (Application.WorksheetFunction.Max(Range(Sheets("Info").Cells(4, 9), Sheets("Info").Cells(va, 13)))) * 1.01
        
    
    '--------Atualiza Gráf 1 Rot Máx / Rot Min / Rot Méd
    
    ActiveSheet.ChartObjects("Gráfico 1").Activate
 
    ActiveChart.FullSeriesCollection(1).Values = "=Info!$C$4:$C$" & va & ""
    ActiveChart.FullSeriesCollection(2).Values = "=Info!$D$4:$D$" & va & ""
    ActiveChart.FullSeriesCollection(3).Values = "=Info!$E$4:$E$" & va & ""
    
    gg = 1
    Do While gg < 4
    ActiveChart.FullSeriesCollection(gg).XValues = "=Info!$B$4:$B$" & va & ""
    gg = gg + 1
    Loop
    
    
    '----------------Atualiza Gráf 3 Range / VA ext
    
    
     ActiveSheet.ChartObjects("Gráfico 3").Activate
 
    ActiveChart.FullSeriesCollection(1).Values = "=Info!$P$4:$P$" & va & ""
    ActiveChart.FullSeriesCollection(2).Values = "=Info!$O$4:$O$" & va & ""
    
    
    gg = 1
    Do While gg < 2
    ActiveChart.FullSeriesCollection(gg).XValues = "=Info!$B$4:$B$" & va & ""
    gg = gg + 1
    Loop
    
    
    '-----------------
    
    Sheets("Info").Activate
    
    Range("N" & va & "").FormatConditions.Add Type:=xlExpression, Formula1:="=N" & va & ">M" & va - 1 & ""
    Range("N" & va & "").FormatConditions(Range("N" & va & "").FormatConditions.Count).SetFirstPriority
    With Range("N" & va & "").FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16711681
        .TintAndShade = 0
    End With
    With Range("N" & va & "").FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 12611584
        .TintAndShade = 0
    End With
    Range("N" & va & "").FormatConditions(1).StopIfTrue = False
    Range("N" & va & "").FormatConditions.Add Type:=xlExpression, Formula1:="=N" & va & "<I" & va - 1 & ""
    Range("N" & va & "").FormatConditions(Range("N" & va & "").FormatConditions.Count).SetFirstPriority
    With Range("N" & va & "").FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Range("N" & va & "").FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
   Range("N" & va & "").FormatConditions(1).StopIfTrue = False
    
    
    Range("A" & va & "").Select
    Range("A1").Select
    
End Sub

Sub EscalaLocal()

'---- acertar valores conforme o filtro
    
    '----
    
    
    tvt1 = 5
    
   Do While Sheets("Comp").Cells(tvt1, 2) <> ""
   tvt1 = tvt1 + 1
   Loop
   tvt1 = tvt1 - 1
    
   Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(tvt1, 3)).FormulaR1C1 = "=RC[-1]"
   Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(tvt1, 4)).FormulaR1C1 = "=RC[-3]"
    
    
    
   tvt = 5
    
   Do While Sheets("Comp").Cells(tvt, 7) <> ""
   tvt = tvt + 1
   Loop
   tvt = tvt - 1
    
   Range(Sheets("Comp").Cells(5, 9), Sheets("Comp").Cells(tvt, 11)).FormulaR1C1 = "=RC[-2]"
   
    
    
    
    '----
    
    gm1 = 4
    
    Do While Sheets("Comp").Cells(gm1, 1) <> ""
    gm1 = gm1 + 1
    Loop
    gm1 = gm1 - 1
    
    
    gm2 = 4
    
    Do While Sheets("Comp").Cells(gm2, 7) <> ""
    gm2 = gm2 + 1
    Loop
    gm2 = gm2 - 1
    
    
    max1 = Application.WorksheetFunction.Max(Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4)))
    min1 = Application.WorksheetFunction.Min(Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4)))
    
    
    max2 = Application.WorksheetFunction.Max(Range(Sheets("Comp").Cells(5, 10), Sheets("Comp").Cells(gm2, 10)))
    min2 = Application.WorksheetFunction.Min(Range(Sheets("Comp").Cells(5, 10), Sheets("Comp").Cells(gm2, 10)))
    
    
    If max1 > max2 Then maxgm = max1 Else maxgm = max2
    If min1 < min2 Then mingm = min1 Else mingm = min2
    
    
    Sheets("Graf").Activate
    
    Range("X12") = maxgm
    Range("X13") = mingm
    Range("X14") = maxgm - mingm
    Range("X15") = (maxgm + mingm) / 2
    Range("X16") = (Application.WorksheetFunction.SumProduct(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)), (Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4))))) / (Application.Sum(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3))))
    
    
   ' k1 = Application.WorksheetFunction.SumProduct(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)), (Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4))))
   ' k2 = Application.Sum(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)))
   ' k3 = k1 / k2
    
    
    desvpv = WorksheetFunction.StDev_P(Range(Sheets("Comp").Cells(5, 4), Sheets("Comp").Cells(gm1, 4)))

    
    vt = Application.Sum(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)))
    vva = vt * 0.707
    
    Range("X18") = desvpv
    
    vp = Application.WorksheetFunction.Max(Range(Sheets("Comp").Cells(5, 3), Sheets("Comp").Cells(gm1, 3)))
    
    lvp = 5
    Do While Sheets("Comp").Cells(lvp, 3) <> vp
    lvp = lvp + 1
    Loop

    POC = Sheets("Comp").Cells(lvp, 4)
    
    Range("X8") = POC
    
    
    lval = lvp - 1
    lvah = lvp + 1

    vcv = vp

    Do While vcv < vva

    vcv = vp + WorksheetFunction.Sum(Range(Sheets("Comp").Cells(lvp - 1, 3), Sheets("Comp").Cells(lval, 3))) + WorksheetFunction.Sum(Range(Sheets("Comp").Cells(lvp + 1, 3), Sheets("Comp").Cells(lvah, 3)))

    If Sheets("Comp").Cells(lval - 2, 3) <> "" Then
    lval = lval - 1
    End If

    If Sheets("Comp").Cells(lvah + 1, 1) <> "" Then
    lvah = lvah + 1
    End If

    Loop


    lval = lval + 1
    lvah = lvah - 1
    
    If lval = 5 Then lval = 6

    VALOW = Sheets("Comp").Cells(lval, 4)
    VAHI = Sheets("Comp").Cells(lvah, 4)
    
    
    Range("X7") = VAHI
    Range("X9") = VALOW
    
    Range("X20") = Sheets("Comp").Cells(5, 10)
    Range("X21") = Sheets("Comp").Cells(gm2, 10)
    
   tvt = 5
    
   Do While Sheets("Comp").Cells(tvt, 12) <> ""
   tvt = tvt + 1
   Loop
   tvt = tvt - 1
   
   Range(Sheets("Comp").Cells(5, 12), Sheets("Comp").Cells(tvt, 14)) = ""
 
    
    Range("Comp!$L$5:$L$" & gm2 & "") = VALOW
    Range("Comp!$M$5:$M$" & gm2 & "") = POC
    Range("Comp!$N$5:$N$" & gm2 & "") = VAHI
    
    Sheets("Graf").Select
   
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(2).Values = "=Comp!$C$5:$C$" & gm1 & ""
    ActiveChart.SeriesCollection(2).XValues = "=Comp!$D$5:$D$" & gm1 & ""
    ActiveChart.SeriesCollection(1).Values = "=Comp!$J$5:$J$" & gm2 & ""
    ActiveChart.SeriesCollection(1).XValues = "=Comp!$G$5:$G$" & gm2 & ""
    ActiveChart.SeriesCollection(3).Values = "=Comp!$N$5:$N$" & gm2 & ""
    ActiveChart.SeriesCollection(3).XValues = "=Comp!$G$5:$G$" & gm2 & ""
    ActiveChart.SeriesCollection(4).Values = "=Comp!$M$5:$M$" & gm2 & ""
    ActiveChart.SeriesCollection(4).XValues = "=Comp!$G$5:$G$" & gm2 & ""
    ActiveChart.SeriesCollection(5).Values = "=Comp!$L$5:$L$" & gm2 & ""
    ActiveChart.SeriesCollection(5).XValues = "=Comp!$G$5:$G$" & gm2 & ""
    
    
   '---- acertar valores da escala
    
    ActiveChart.Axes(xlValue).Select
    
    ActiveChart.Axes(xlValue).MaximumScale = maxgm
    ActiveChart.Axes(xlValue).MinimumScale = mingm
    

    ActiveChart.ChartArea.Select
    Range("A1").Select
    


End Sub
