Attribute VB_Name = "Module1"
' constantes
Private Const sheetWIR As String = "liste", sheetModel As String = "model", sheetHisto As String = "histo"
Private Const tblOfPoints As String = "t_RLA", tblOfPointsModel As String = "T_model", tblHistorical As String = "T_histo"
Private Const fileNumber As String = "1", addressNbrOfLUP As String = "C3"
Private Const version As String = "23.3", versionDate As Date = "02/08/2023"
'
Type photoTemplate
    insertPoint As Long
    progressValue As Single
    bldName As String
    photoSheet As String
    LUPsheet As String
    commentPoint As String
    imageName As String
    imageFile As String
End Type
'
Public Type histoTable
    dateSend As Date
    typeSend As String
    senderName As String
    fName As String
    passWord As String
    nCreation As Long
    nEnvoye As Long
    nEncours As Long
    nEnAttente As Long
    nEncoursBack As Long
    nAnnuleBack As Long
    nSoldeeBack As Long
End Type
Sub Versionning()
    Call PutValueOnTable_Data("T_parameters", 2, "version", version)
    Call PutValueOnTable_Data("T_parameters", 2, "versionDate", versionDate)
End Sub
Sub InsertNewRows()
Dim nOfLUP As Long, curTblRow As Long, nTblCol As Long, dateCol As Long
Dim varName As String
Dim tablePS As ListObject
'Set path for Table variable
Set tablePS = ActiveSheet.ListObjects(tblOfPoints)
'
    Application.EnableEvents = False
'
    nOfLUP = Range(addressNbrOfLUP).Value
    curTblRow = nOfLUP + GetValueFromTable_Data("T_parameters", 2, "firstRow") - 1 - 11
'
' Unprotecting current sheet
'
    ActiveSheet.Unprotect
'
' Insert new rows in the liste and populate columns "date saisie" and "type"
'
    tablePS.ListRows.Add AlwaysInsert:=True
'
    dateCol = GetColNumberFromName(tblOfPoints, "date ouverture")
    curTblRow = curTblRow + 1
'
    Call PutValueOnTable_ColRow(tablePS, 1, curTblRow, curTblRow)
    Call PutValueOnTable_ColRow(tablePS, dateCol, curTblRow, Now())
'
    Call PutValueOnTable_ColRow(tablePS, GetColNumberFromName(tblOfPoints, "statut"), curTblRow, "CREATION")
'    Call PutValueOnTable_ColRow(tablePS, GetColNumberFromName(tblOfPoints, "statut EXE"), curTblRow, "A faire")
'
' Range(Cells(curRow, 16), Cells(curRow, 16)).Formula = "=CONCATENATE(data!$P$1,data!$P$2," & ActiveSheet.Name & "!O" & curRow & " )"
' Unprotect cells on current row
'    Range(Cells(curRow, 3), Cells(curRow, 9)).Locked = False
'    'Range("Z1").Value = Range("Z1").Value + 1
'   Protecting again the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True
'
Range(Cells(curTblRow + GetValueFromTable_Data("T_parameters", 2, "firstRow") - 1, 1), _
      Cells(curTblRow + GetValueFromTable_Data("T_parameters", 2, "firstRow") - 1, 1)).Select
'
Application.EnableEvents = True
'
Set tablePS = Nothing

'
End Sub
Sub DeleteLastRow()
Dim lstRow As Long, nOfLUP As Long, iCol1 As Long, icol2 As Long
    nOfLUP = Range(addressNbrOfLUP).Value
    If nOfLUP = 1 Then
        MsgBox ("La liste contient une seule ligne")
        Exit Sub
    End If
'
    Application.EnableEvents = False
'
    lstRow = nOfLUP + GetValueFromTable_Data("T_parameters", 2, "firstRow") - 1
'
' Delete last line
'
' Verify if this line has statut = "CREATION"
'
    iCol1 = GetColNumberFromName(tblOfPoints, "statut")
    If Range(Cells(lstRow, iCol1), Cells(lstRow, iCol1)).Value = "CREATION" Then
        ActiveSheet.Unprotect
        Rows(lstRow & ":" & lstRow).Select
        Selection.Delete Shift:=xlUp
        
    Else
        MsgBox ("le statut de la dernière ligne est différent de 'CREATION'")
    
    End If
'
    Application.EnableEvents = True
'
' Protect the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True
End Sub
Sub PrintSheets()
Dim sh As Worksheet, nOfLUP As Long, lstRow As Long
'
    nOfLUP = Range(addressNbrOfLUP).Value
'
    lstRow = nOfLUP + GetValueFromTable_Data("T_parameters", 2, "firstRow") - 1
' Scan over tabs
    For Each sh In ActiveWorkbook.Worksheets
        Select Case sh.Name
        Case "data"
        
        Case Else
            ActiveSheet.PageSetup.PrintArea = "$A$1:$O$" & lstRow
            ActiveSheet.PageSetup.CenterHeader = "&B <<TBS4 - Liste de Points Ouverts>>"
            sh.PrintOut Preview:=True
        
        End Select
    
    Next sh
    ActiveWorkbook.Protect Structure:=True, Windows:=False
'
End Sub
Sub PrintActiveSheet()
'
    ActiveSheet.PrintOut Preview:=True
'
End Sub
Sub PrintAsPDF()
Dim fName As String, ExcelVer As String, curPhoto As photoTemplate, iRow As Long, iPhotos As Long, nPhotos As Long
Dim lstRow As Long, nOfLUP As Long, seqData As String
'
    Application.EnableEvents = False
'
    nOfLUP = Range(addressNbrOfLUP).Value
'
    lstRow = nOfLUP + GetValueFromTable_Data("T_parameters", 2, "firstRow") - 1
'   defintion of the new table name and new sheet name
    seqDate = Year(Date) - 2000 & Format(Month(Date), "00") & Format(Day(Date), "00")
'
    ExcelVer = Application.version
'
    Select Case ExcelVer
    Case Is <> "15.0"
        i = MsgBox("Cette version d'Excel n'accepte pas le PDF direct. Veuillez utiliser la commande 'impression' et PDF Creator", vbOKOnly)
        
    Case Else
        'iReturn = MsgBox("Incluir além da lista também um anexo com fotos?", vbQuestion + vbYesNo, "imprir o anexo")
        iReturn = vbNo ' uniquement liste sans photo pour l'instant
        Select Case iReturn
        Case vbNo
'       list only - part 1
            fName = ActiveWorkbook.Path & "\pdf\" & seqDate & "_liste_" & ActiveWorkbook.Name & ".pdf"
            ActiveSheet.PageSetup.PrintArea = "$A$1:$O$" & lstRow
            ActiveSheet.PageSetup.CenterHeader = "&B <<TBS4 - Liste de Points Ouverts>>"
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fName, Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
'       list only - part 2
'            ActiveSheet.Unprotect
'            Range("I:R").EntireColumn.Hidden = True
'            fName = ActiveWorkbook.Path & "\pdf\" & seqDate & "_horsBAG-part2_" & ActiveWorkbook.Name & ".pdf"
'            ActiveSheet.PageSetup.PrintArea = "$A$1:$U$" & lstRow
'            ActiveSheet.PageSetup.CenterHeader = "&B <<Part 2>>"
'            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fName, Quality:=xlQualityStandard, _
'                IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
'            Range("I:R").EntireColumn.Hidden = False
'            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
'        , AllowFiltering:=True
        
        Case vbYes
'       list + appendix photos
'           counting number of photos
            nPhotos = 0
            For iRow = 8 To Range("I4").Value + 7
                If Cells(iRow, 15) <> "" Then nPhotos = nPhotos + 1
                
            Next iRow
            If nPhotos > 0 Then
'           photos
                curPhoto.bldName = Sheets("Sumário").Range("E3").Value
                curPhoto.LUPsheet = ActiveSheet.Name
                curPhoto.photoSheet = "photoAll"
'           create a temporary sheet
                Application.ScreenUpdating = False
'           Use the Status Bar to inform  user of the macro's progress
'           change the cursor to hourglass
                Application.Cursor = xlWait
'           makes sure that the statusbar is visible
                Application.DisplayStatusBar = True
'           add your message to status bar
                Application.StatusBar = "Criando arquivos PDF..."
                
                ActiveWorkbook.Unprotect
                Sheets("model").Visible = True
                Sheets("model").Select
                Sheets("model").Copy After:=Sheets(Sheets.Count)
                Sheets("model (2)").Select
                Sheets("model (2)").Name = curPhoto.photoSheet
'
                curPhoto.insertPoint = 1
                For iPhotos = 1 To nPhotos
                    curPhoto.insertPoint = copyTemplatePhoto(curPhoto)
                
                Next iPhotos
'
                curPhoto.insertPoint = 1
                For iRow = 8 To Sheets(curPhoto.LUPsheet).Range("I4").Value + 7
                    If Sheets(curPhoto.LUPsheet).Cells(iRow, 15).Value <> "" Then
'
                        Sheets(curPhoto.LUPsheet).Activate
                        Cells(iRow, 1).Activate
                        Call setHeading(ActiveCell.Row, curPhoto)
'
                        curPhoto.insertPoint = postPhoto(curPhoto)
'
                    End If
'
                Next iRow
'
'               print list
                Sheets(curPhoto.LUPsheet).Activate
                fName = ActiveWorkbook.Path & "\pdf\" & ActiveWorkbook.Name & "_" & ActiveSheet.Name & ".pdf"
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fName, Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
'               print photos
                Sheets(curPhoto.photoSheet).Activate
                fName = ActiveWorkbook.Path & "\pdf\" & ActiveWorkbook.Name & "_" & curPhoto.LUPsheet & "_fotos.pdf"
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fName, Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
'
                Application.DisplayAlerts = False
                ActiveWindow.SelectedSheets.Delete
                Application.DisplayAlerts = True
                Sheets("model").Visible = False
'
                ActiveWorkbook.Protect Structure:=True, Windows:=False
'
                Sheets(curPhoto.LUPsheet).Select
'               restore default cursor
                 Application.Cursor = xlDefault
'               gives control of the statusbar back to the programme
                Application.StatusBar = False
'
                Application.ScreenUpdating = True
                
            Else
                fName = ActiveWorkbook.Path & "\pdf\" & ActiveWorkbook.Name & "_" & ActiveSheet.Name & ".pdf"
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fName, Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

            End If
        End Select
    End Select
'
    Application.EnableEvents = True
'
End Sub
Sub consolidationWorkbook()
'
'   Consolidation of a sent list
'
'   Objective: procedure activated by a button located on Histo tab
'   Instructions: select file to consolidate on the list and push button.
'
    Dim tableHisto As ListObject, tableConso As ListObject, tableWIR As ListObject
    Dim rowHisto As histoTable, dateConsolidation As Date
    Dim curTblRow As Long, iCol As Long
    Dim kLine As Long, sLine As Long, nbrLines2conso As Long, nbrOfLUP As Long
    Dim fNameConso As String, fNameMain As String, consoSheet As String, fullNameConso As String, lineID As String, passWord As String
    Dim SelectedCell As Range
'   Verify if selected line is inside T_histo table
'
    If Not DetermineActiveTable(tblHistorical) Then
        MsgBox ("Sélection érronée. Veuillez placer le curseur sur la table des historiques")
        Exit Sub
    
    End If
    'Set SelectedCell = ActiveCell
'   assignment for table T_histo
    Set tableHisto = Sheets(sheetHisto).ListObjects(tblHistorical)
    curTblRow = ActiveCell.Row - 4 'row number inside of table
'   Verify if the selected line a "envoyé" file line
    If Not (GetValueFromTable_ColRow(tableHisto, 13, curTblRow) = "GO") Then
        MsgBox ("Sélection érronée. Veuillez placer le curseur sur une ligne d'un fichier envoyé (statut : GO)")
        Set tableHisto = Nothing
'
        Exit Sub
    End If
'   retreive actual workbook name
    fNameMain = Application.ActiveWorkbook.Name
'   set table of points
    Set tableWIR = Sheets(sheetWIR).ListObjects(tblOfPoints)
'   unprotect sheetWIR
    Worksheets(sheetWIR).Unprotect
'   get number of points (nbrOfUP)
    nbrOfLUP = Sheets(sheetWIR).Range("O7").Value
'
'   retrieve fileName
    fNameConso = GetValueFromTable_ColRow(tableHisto, 4, curTblRow) & ".xlsx"
    fullNameConso = Application.ActiveWorkbook.Path & "\Reçus\LPS-002\" & fNameConso
'
    Workbooks.Open (fullNameConso)
'
    consoSheet = "L_" & Mid(fNameConso, 1, 6)
    Sheets(consoSheet).Activate
    Set tableConso = Sheets(consoSheet).ListObjects("T_" & Mid(fNameConso, 1, 6))
'
'   =========================================================
'   running of consolidation
'   For loop over lines to consolidate
    nbrLines2conso = Range("A8").Value
    dateConsolidation = Now()
    passWord = GetValueFromTable_ColRow(tableHisto, 5, curTblRow)
'
    For kLine = 1 To nbrLines2conso
        lineID = GetValueFromTable_ColRow(tableConso, 2, kLine)
        sLine = 1
        Do
            'If sLine > nbrOfLUP Then Exit Do
            If lineID = GetValueFromTable_ColRow(tableWIR, 2, sLine) Then
                For iCol = 18 To 22
'                   column 18 = "état CNS"
'                   column 19 = "% Avancement"
                    Call PutValueOnTable_ColRow(tableWIR, iCol, sLine, _
                         GetValueFromTable_ColRow(tableConso, iCol, kLine) _
                         )
'
                Next iCol
'               column 29 = "consolid_ID"
                Call PutValueOnTable_ColRow(tableWIR, 29, sLine, passWord)
'               column 30 = "dateConsol"
                Call PutValueOnTable_ColRow(tableWIR, 30, sLine, dateConsolidation)
'
            End If
            sLine = sLine + 1
        Loop Until sLine > nbrOfLUP
    Next kLine
'   =========================================================
    consoSheet = "L_" & Mid(fNameConso, 1, 6)
    Sheets(consoSheet).Activate
'   retrieve valeurs on heading on external file (fullNameConso file)
    rowHisto.nAnnuleBack = Range("R2").Value
    rowHisto.nEncoursBack = Range("R3").Value
    rowHisto.nSoldeeBack = Range("R6").Value
'   close workbook fNameConso
    Workbooks(fNameConso).Close
'   re-activate histo table
    Workbooks(fNameMain).Activate
'
    Call fillTableHisto("BACK", tableHisto, rowHisto, curTblRow)
'
'
    Worksheets(sheetHisto).Protect
'
    Worksheets(sheetWIR).Activate
    Worksheets(sheetWIR).Protect
'
    Set tableHisto = Nothing
'
End Sub
Sub sendModifiedLines()
'
'   Creation of a file with newly modified lines only
'
Dim typeSend As String, dateSend As Date
'   Test if new lines exist
    If Range("Y8").Value = 0 Then
        MsgBox ("Aucune ligne récemment modifée. Pas de transfert")
        Exit Sub
    
    End If
'
    dateSend = Date
    Call sendLines("modified", dateSend)
'
End Sub
Sub sendCreatedLines()
'
'   Creation of a file with newly created lines only
'
Dim typeSend As String, dateSend As Date
'   Test if new lines exist
    If Range("Q8").Value = 0 Then
        MsgBox ("Aucune ligne récemment créée. Pas de transfert")
        Exit Sub
    
    End If
'   Test if all new lines have date on field "date envoi"
    dateSend = Range("M9").Value
    If dateSend = 0 Then
        MsgBox ("Il existe au moins une ligne nouvelle sans data d'envoi. Pas de transfert")
        Exit Sub
    
    End If
    Call sendLines("new", dateSend)
'
End Sub
Sub sendAllLines()
'
'   Creation of a file with all lines except "reçu CNS", "soldée" and "annulé"
'
Dim typeSend As String, dateSend As Date
'
    dateSend = Date
    Call sendLines("new+old", dateSend)
'
End Sub
Sub sendLines(typeSend As String, dateSend As Date)
'
Dim nameNewSheet As String, seqDate As String, nameNewFile As String
Dim fPath As String, fName As String, myPassword As String
Dim SelectedCell As Range, tableWIR As ListObject, tableModel As ListObject, tableHisto As ListObject
Dim destinationRow As ListRow, originRow As Range
Dim iRow As Long, nCol As Long, nColFilter As Long, curTblRow As Long
Dim arrList() As String, rowHisto As histoTable
'
'   determination of column number for filtering
'   "état MOE" on tableWIR for P and T
'   "modifé?" on table WIR for M
'
'   unprotect sheetWIR and sheetModel
    Worksheets(sheetWIR).Unprotect
    Worksheets(sheetModel).Unprotect
'   assignments for tables on Wir (origine) and on model (destination)
    Set tableWIR = Sheets(sheetWIR).ListObjects(tblOfPoints)
    Set tableModel = Sheets(sheetModel).ListObjects(tblOfPointsModel)
    Set tableHisto = Sheets(sheetHisto).ListObjects(tblHistorical)
'
    Select Case typeSend
    Case "new"
        ReDim Preserve arrList(0)
        arrList(0) = "création"
'
        appFname = "_P_"
        nColFilter = GetColNumberFromTable(tableWIR, "état MOE")
        rowHisto.typeSend = "création"
        
    Case "modified"
        ReDim Preserve arrList(0 To 1)
        arrList(0) = "O"
        arrList(1) = "o"
'
        appFname = "_M_"
        nColFilter = GetColNumberFromTable(tableWIR, "modifié?")
        rowHisto.typeSend = "modification"
        
    Case "new+old"
        ReDim Preserve arrList(0 To 3)
        arrList(0) = "création"
        arrList(1) = "envoyé"
        arrList(2) = "en cours"
        arrList(3) = "En attente"
'
        appFname = "_T_"
        nColFilter = GetColNumberFromTable(tableWIR, "état MOE")
        rowHisto.typeSend = "total"
    
    End Select
'   unprotect sheetWIR and sheetModel
    Worksheets(sheetWIR).Unprotect
    Worksheets(sheetModel).Unprotect
'   remove all filters
    tableWIR.AutoFilter.ShowAllData
'   filter lines with état MOE (case P or T) or modifié? (case M) following arrLst contents
    With tableWIR
        .Range.AutoFilter Field:=nColFilter, Criteria1:=arrList, Operator:=xlFilterValues
        SourceDataRowsCount = .ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible).Count
        .DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
    
    End With
'   Copy only visible cells on tableModel
    With tableModel
        '.Resize .Range.Resize(.Range.Rows.Count + SourceDataRowsCount, .Range.Columns.Count)
        .Resize .Range.Resize(1 + SourceDataRowsCount, .Range.Columns.Count)
        .DataBodyRange.Cells(1, 1).PasteSpecial (xlPasteValues)
    
    End With
    Application.CutCopyMode = False
'   remove all filters from tableWIR
    tableWIR.AutoFilter.ShowAllData
'
    
    If (Mid(typeSend, 1, 3) = "new") Then
'       for loop over tableWIR - searching of "création" lines to change value for "envoyé" (valable also for "new+old" case)
        For iRow = 1 To tableWIR.ListRows.Count
            If (GetValueFromTable_ColRow(tableWIR, nColFilter, iRow) = "création") Then
                Call PutValueOnTable_ColRow(tableWIR, nColFilter, iRow, "envoyé")
        
            End If
        
        Next iRow
        
    ElseIf (typeSend = "modified") Then
'       for loop over tableWIR - searching of "O" lines (on modifié?) to change value for none
        For iRow = 1 To tableWIR.ListRows.Count
            If (GetValueFromTable_ColRow(tableWIR, nColFilter, iRow) = "o" Or _
                GetValueFromTable_ColRow(tableWIR, nColFilter, iRow) = "O") Then
                Call PutValueOnTable_ColRow(tableWIR, nColFilter, iRow, "")
           
            End If
        
        Next iRow
        
    End If
'   defintion of the new table name and new sheet name
    seqDate = Year(dateSend) - 2000 & Format(Month(dateSend), "00") & Format(Day(dateSend), "00")
'
    nameNewSheet = "L_" & seqDate
    nameNewTable = "T_" & Mid(nameNewSheet, 3)
    fPath = ThisWorkbook.Path & "\envois\LPS-099\"
    fName = seqDate & appFname & "4411550-Q-PAQ-LPS-0099"
    myPassword = uniqueCode(fName, Now, randLong(Now))
'   select model sheet to copy on a new workbook
    Sheets(sheetModel).Activate
    If copySheetOnNewFile(typeSend, sheetModel, fPath, fName) Then
        Sheets(sheetModel).Range("B1").Value = typeSend
'
        Sheets(sheetModel).Range("R2").Formula = "=COUNTIF(T_model[état CNS],$P2)"
        Sheets(sheetModel).Range("R3").Formula = "=COUNTIF(T_model[état CNS],$P3)"
        Sheets(sheetModel).Range("R6").Formula = "=COUNTIF(T_model[état CNS],$P6)"
'       number of points per type
        rowHisto.nEncours = Sheets(sheetModel).Range("Q3").Value
        rowHisto.nEnvoye = Sheets(sheetModel).Range("Q4").Value
        rowHisto.nEnAttente = Sheets(sheetModel).Range("Q5").Value
        rowHisto.nCreation = Sheets(sheetModel).Range("Q8").Value
'       wrap text for all columns on the table
        Sheets(sheetModel).ListObjects(1).DataBodyRange.WrapText = True
'       reanme table
        Sheets(sheetModel).ListObjects(1).Name = nameNewTable
'       protect sheet with a password
        Sheets(sheetModel).Protect passWord:=myPassword, Contents:=True, Scenarios:=True, _
                                  AllowFiltering:=True, DrawingObjects:=True
'       show sheetName before close file
        Sheets(sheetModel).Visible = True
'       reanme sheet
        Sheets(sheetModel).Name = nameNewSheet
'       close file
        ActiveWorkbook.Close SaveChanges:=True
'       fill up table histoTable
        Sheets(sheetHisto).Activate
'
        rowHisto.dateSend = dateSend
        rowHisto.senderName = Environ("username")
        rowHisto.fName = fName
        rowHisto.passWord = myPassword
'
        curTblRow = Range("A3").Value + 1
'
        Call fillTableHisto("GO", tableHisto, rowHisto, curTblRow)
'       .....
        
    End If
'   clean up contents of table modelTable
    tableModel.DataBodyRange.Delete
'   unassigments
    Sheets(sheetWIR).Activate
    Sheets(sheetWIR).Protect
'
    Set tableWIR = Nothing
    Set tableModel = Nothing
    Set tableHisto = Nothing
'
End Sub
Sub fillTableHisto(way As String, tableHisto As ListObject, rowHisto As histoTable, curTblRow As Long)
'dateListe As Date, typeSend As String, n1 As Long, n2 As Long, n3 As Long, n4 As Long, fName As String, passWord As String)
'
' table:
'   1 - date liste
'   2 - type envoi
'   3 - expéditeur
'   4 - nom fichier
'   5 - mot de passe
'   6 - nbr création (GO)
'   7 - nbr envoyé (GO)
'   8 - nbr en cours (GO)
'   9 - nbr En attente (GO)
'  10 - nbr en cours (BACK)
'  11 - nbr annulé (BACK)
'  12 - nbr soldée (BACK)
'  13 - statut (GO ou BACK)
'  14 - date de la consolidation
'
'
'   Unprotecting the sheet
    ActiveSheet.Unprotect

Select Case way
Case "GO"
'   sending data
'   Insert new rows in the liste and populate columns
    tableHisto.ListRows.Add AlwaysInsert:=True
'
    tableHisto.DataBodyRange.Cells(curTblRow, 1).Value = rowHisto.dateSend
    tableHisto.DataBodyRange.Cells(curTblRow, 2).Value = rowHisto.typeSend
    tableHisto.DataBodyRange.Cells(curTblRow, 3).Value = rowHisto.senderName
    tableHisto.DataBodyRange.Cells(curTblRow, 4).Value = rowHisto.fName
    tableHisto.DataBodyRange.Cells(curTblRow, 5).Value = rowHisto.passWord
    tableHisto.DataBodyRange.Cells(curTblRow, 6).Value = rowHisto.nCreation
    tableHisto.DataBodyRange.Cells(curTblRow, 7).Value = rowHisto.nEnvoye
    tableHisto.DataBodyRange.Cells(curTblRow, 8).Value = rowHisto.nEncours
    tableHisto.DataBodyRange.Cells(curTblRow, 9).Value = rowHisto.nEnAttente
    tableHisto.DataBodyRange.Cells(curTblRow, 13).Value = "GO"

Case "BACK"
'   receiveing data
'
    tableHisto.DataBodyRange.Cells(curTblRow, 10).Value = rowHisto.nEncoursBack
    tableHisto.DataBodyRange.Cells(curTblRow, 11).Value = rowHisto.nAnnuleBack
    tableHisto.DataBodyRange.Cells(curTblRow, 12).Value = rowHisto.nSoldeeBack
    tableHisto.DataBodyRange.Cells(curTblRow, 13).Value = "BACK"
    tableHisto.DataBodyRange.Cells(curTblRow, 14).Value = Now()
'

End Select
'   Protecting again the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True

End Sub
Sub insertNewSubject()
Dim ws As Worksheet, tbl As ListObject, newRow As ListRow, sortcolumn As Range, newValue As Range
    Set ws = Sheets("data")
    Set newValue = Range("G9")
'
    If (newValue.Value = "" Or ws.Range("N28").Value > 0) Then Exit Sub
'
    ws.Unprotect
    Set tbl = ws.ListObjects("t_thèmes")
'   adding a new row after last row
    Set newRow = tbl.ListRows.Add
'   populate field
    With newRow
        .Range(1) = newValue.Value
    End With
'   sorting table
    Set sortcolumn = Range("t_thèmes[thème]")
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
'   Clearing field
    newValue.Value = ""
'   Protecting again the sheet
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True
'
'   Inclusion of a new line in t_parSujet
    Set ws = Nothing
    Set tbl = Nothing
    Set newRow = Nothing
    Set newValue = Nothing
'
End Sub
Sub insertNewCompany()
Dim ws As Worksheet, tbl As ListObject, newRow As ListRow, sortcolumn As Range, newValue As Range
    Set ws = Sheets("data")
    Set newValue = Range("M9")
'
    If (newValue.Value = "" Or ws.Range("C2").Value > 0) Then Exit Sub
'
    ws.Unprotect
    Set tbl = ws.ListObjects("t_entreprises")
'   adding a new row after last row
    Set newRow = tbl.ListRows.Add
'   populate field
    With newRow
        .Range(1) = newValue.Value
    End With
'   sorting table
    Set sortcolumn = Range("t_entreprises[entreprise]")
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
'   Clearing field
    newValue.Value = ""
'   Protecting again the sheet
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True
'
'   Inclusion of a new line in t_parSujet
    Set ws = Nothing
    Set tbl = Nothing
    Set newRow = Nothing
    Set newValue = Nothing
'
End Sub
Sub insertOLE()
Dim myLink
'   Verify if selected line is inside 'table of points'
'
    Application.EnableEvents = False
'
    If Application.ActiveCell.Row < 12 Or _
       Not Application.ActiveCell.Column = 16 Then
        MsgBox ("Sélection érronée. Veuillez placer le curseur sur une cellule de la colonne 'OLE'")
        Application.EnableEvents = True
        Exit Sub
        
    End If
'   Unprotecting the sheet
    ActiveSheet.Unprotect

'    myLink = Application.Dialogs(xlDialoglink).Show
'   another away
    
    myLink = Application.GetOpenFilename()
    If myLink = "False" Then Exit Sub
    If ActiveCell.Value = "" Then ActiveCell.Value = "link"
    ActiveCell.Hyperlinks.Add ActiveCell, myLink, , , ActiveCell.Value
'   end of another way
'   Protecting again the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True
    
    Application.EnableEvents = True
    
    
End Sub
Sub insertRemarksFromMOE()
'
Dim tri As String
    
'   test if active cursor is in column 'commentaires AMOE/MOE'
    If (ActiveCell.Column = 15) And (ActiveCell.Row > 11) Then
'
        tri = GetValueFromTable_Data("t_parameters", 2, "currentTri")
'
        insertRemarks.TextBox1.Value = ActiveCell.Value & _
                                        vbLf & Format(Now(), "dd/mm") & " [" & tri & "] "
        
'
        insertRemarks.Show
'
    End If
'
End Sub
Sub UpdateFunctions()
    Worksheets("liste").Range("N8").Calculate
    
End Sub
Sub insertHyperLink()
Dim rng As Range, iStart As Long
Dim tableWIR As ListObject, iRow As Long
Dim x As Long, baloonStr As String
'   Verify if selected line is inside T_histo table
'
    If Not DetermineActiveTable(tblOfPoints) Then
        MsgBox ("Sélection érronée. Veuillez placer le curseur sur la table des RPO")
        Exit Sub
    
    End If
    
    Application.EnableEvents = False
 
'Set path for Table variable and active cell
    Set tableWIR = Sheets(sheetWIR).ListObjects(tblOfPoints)
    Set rng = Range(Selection.Address)
    
'Searching for tag ([RPO nnnn])
    iRow = 0
    iStart = positionTag(tableWIR, rng, iRow, baloonStr) ' if iStart is negative means no tag is found or tag is meaningless
    If iStart = 0 Then Exit Sub
    
'   Unprotecting the sheet
    ActiveSheet.Unprotect

    rng.Parent.Hyperlinks.Add _
        Anchor:=rng, _
        Address:="", _
        SubAddress:="liste!B" & iRow, _
        ScreenTip:=baloonStr, _
        TextToDisplay:=rng.Value

    With rng.Font
        .ColorIndex = xlAutomatic
        .Underline = xlUnderlineStyleNone
        .Name = "Calibri"
        .Size = 10
        
    End With

    With rng.Characters(Start:=iStart, Length:=11).Font
        .Underline = xlUnderlineStyleSingle
        .Color = -4165632
    End With
    
    Set tableWIR = Nothing
    Set rng = Nothing

'   Protecting again the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True
'
    Application.EnableEvents = True
'
End Sub
Function positionTag(tbl As ListObject, rng As Range, iRow As Long, baloonStr As String) As Long
Dim iTag As Long, myArray As Variant, x As Long, valRPO As Long
Dim posTag1 As Long, posTag2 As Long
    
    iTag = InStr(1, rng.Value, "[RPO", vbTextCompare)
    If iTag <> 0 Then
        posTag1 = iTag + Len("RPO") + 1
        posTag2 = InStr(posTag1, rng.Value, "]", vbTextCompare)
        If posTag2 = 0 Then
            positionTag = posTag2
            Exit Function
        End If
        valRPO = CLng(Mid(rng.Value, posTag1, posTag2 - posTag1))
        'MsgBox ("1: " & posTag1 & " 2: " & posTag2 & " val: " & valRPO)
        'Exit Function
        'Create Array List from Table
        myArray = tbl.DataBodyRange
        'Looping for matching
        For x = LBound(myArray) To UBound(myArray)
            If myArray(x, 1) = valRPO Then
                baloonStr = myArray(x, 6)
                positionTag = 9999
                Exit For
            
            End If
        
        Next x
        
        If x > UBound(myArray) Then
            positionTag = 0
        
        Else
            iRow = x + GetValueFromTable_Data("T_parameters", 2, "firstRow") - 1
        
        End If
    
    Else
        positionTag = iTag
        
    End If
    
    
End Function
Sub deleteHyperLink()
Dim rng As Range, iStart As Long
Dim tableWIR As ListObject, iRow As Long
Dim x As Long, baloonStr As String
'   Verify if selected line is inside T_histo table
'
    Application.EnableEvents = False
'
    If Not DetermineActiveTable(tblOfPoints) Then
        MsgBox ("Sélection érronée. Veuillez placer le curseur sur la table des RPO")
        Exit Sub
    
    End If

'Set path for Table variable and active cell
    Set tableWIR = Sheets(sheetWIR).ListObjects(tblOfPoints)
    Set rng = Range(Selection.Address)
    
'Searching for tag ([RPO nnn])
    iRow = 0
    iStart = positionTag(tableWIR, rng, iRow, baloonStr) ' if iStart is negative meaning no tag is found or tag is meaningless
    If iStart = 0 Then Exit Sub
    
'   Unprotecting the sheet
    ActiveSheet.Unprotect

    With rng
        .Hyperlinks.Delete
        .Locked = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    
    End With
    
    With rng.Font
        .Name = "Calibri"
        .Size = 10
        
    End With
    
    Set tableWIR = Nothing
    Set rng = Nothing
'
    Application.EnableEvents = True
'   Protecting again the sheet
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True

End Sub

