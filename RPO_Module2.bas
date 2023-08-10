Attribute VB_Name = "Module2"
Public Function GetValueFromTable_ActiveSheet(tblName As String, colNum As Long, varName As String) As Variant
'
Dim myTable As ListObject
Dim myArray As Variant
Dim x As Long
'
On Error GoTo errorHandler
'
'Set path for Table variable
Set myTable = ActiveSheet.ListObjects(tblName)

'Create Array List from Table
myArray = myTable.DataBodyRange
'Looping for matching
For x = LBound(myArray) To UBound(myArray)
        If myArray(x, 1) = varName Then
            GetValueFromTable = myTable.DataBodyRange(x, colNum).Value
            Exit Function
        End If
Next x
Call Err.Raise(vbObjectError + 1010, "GetValueFromTable", "no matching found")

errorHandler:
    MsgBox ("[" & Err.Number - vbObjectError & "] " & Err.Source & ", " & Err.Description)

End Function
Public Sub PutValueOnTable_ActiveSheet(tblName As String, colNum As Long, varName As String, Value As Variant)
'
'On Err GoTo 0
'
Dim myTable As ListObject
Dim myArray As Variant
Dim x As Long
'
'Set path for Table variable
Set myTable = ActiveSheet.ListObjects(tblName)
'Create Array List from Table
myArray = myTable.DataBodyRange
'Looping for matching
For x = LBound(myArray) To UBound(myArray)
        If myArray(x, 1) = varName Then
            myTable.DataBodyRange(x, colNum).Value = Value
            Exit Sub
'
        End If
Next x
'Call Err.Raise(1000, "PutValueOnTable", "no matching found")
 
End Sub
Public Sub PutValueOnTable_Data(ByVal tblName As String, ByVal colNum As Long, ByVal varName As String, ByVal Value As Variant)
'
'On Err GoTo 0
'
Dim myTable As ListObject
Dim myArray As Variant
Dim x As Long
'
'Set path for Table variable
Set myTable = Sheets("data").ListObjects(tblName)
'Create Array List from Table
myArray = myTable.DataBodyRange
'Looping for matching
For x = LBound(myArray) To UBound(myArray)
        If myArray(x, 1) = varName Then
            myTable.DataBodyRange(x, colNum).Value = Value
            Set myTable = Nothing
            Exit Sub
'
        End If
Next x
'Call Err.Raise(1000, "PutValueOnTable", "no matching found")
 
End Sub
Public Function GetValueFromTable_Data(ByVal tblName As String, ByVal colNum As Variant, ByVal varName As String) As Variant
'
Dim myTable As ListObject
Dim myArray As Variant
Dim x As Long
'
'On Error GoTo errorHandler
'
'Set path for Table variable
Set myTable = Sheets("data").ListObjects(tblName)

'Create Array List from Table
myArray = myTable.DataBodyRange
'Looping for matching
For x = LBound(myArray) To UBound(myArray)
        If myArray(x, 1) = varName Then
            GetValueFromTable_Data = myTable.DataBodyRange(x, colNum).Value
            Set myTable = Nothing
            Exit Function
        End If
Next x
'Call Err.Raise(vbObjectError + 1010, "GetValueFromTable", "no matching found")
'GetValueFromTable = vbObjectError + 1010
'GetValueFromTable_Data = "not found"

'errorHandler:
'    MsgBox ("[" & Err.Number - vbObjectError & "] " & Err.Source & ", " & Err.Description)

End Function
Public Function GetValueFromTable(myTable As ListObject, colNum As Variant, varName As String) As Variant
'
Dim myArray As Variant
Dim x As Long
'
'Create Array List from Table
myArray = myTable.DataBodyRange
'Looping for matching
For x = LBound(myArray) To UBound(myArray)
        If myArray(x, 1) = varName Then
            GetValueFromTable = myTable.DataBodyRange(x, colNum).Value
            Exit Function
        End If
Next x
'
GetValueFromTable = "ERR.GetValueFromTable: no matching found"
'
End Function
Public Function GetValueFromTable_ColRow(myTable As ListObject, colNum As Long, rowNum As Long) As Variant
'
If (colNum > myTable.ListColumns.Count) Or (rowNum > myTable.ListRows.Count) Then
    Call Err.Raise(1010, "GetValueFromTable_ColRow", "number of columns or rows exceed")

Else
    GetValueFromTable_ColRow = myTable.DataBodyRange(rowNum, colNum).Value
    
End If
'
End Function

Public Sub PutValueOnTable_ColRow(ByVal myTable As ListObject, ByVal colNum As Long, ByVal rowNum As Long, ByVal Value As Variant)
'
If (colNum > myTable.ListColumns.Count) Or (rowNum > myTable.ListRows.Count) Then
    Call Err.Raise(1010, "PutValueOnTable_ColRow", "number of columns or rows exceed")

Else
    myTable.DataBodyRange(rowNum, colNum).Value = Value
    
End If
'
End Sub
Public Function GetColNumberFromTable(myTable As ListObject, colName As String) As Long
' avant l'appel faire les instrution :
'           Dim myTable as ListObject
'           Set myTable = Sheets("sheet name").ListObjects(tblName)
'
Dim x As Long
'
For x = 1 To myTable.ListColumns.Count
    If colName = myTable.ListColumns(x).Name Then
        GetColNumberFromTable = x
        Exit Function
        
    End If
'
Next x
End Function
'
Public Function GetColNumberFromName(tblName As String, colName As String) As Long
'
Dim myTable As ListObject
Dim x As Long
'Set path for Table variable
Set myTable = ActiveSheet.ListObjects(tblName)

For x = 1 To myTable.ListColumns.Count
    If colName = myTable.ListColumns(x).Name Then
        GetColNumberFromName = x
        Exit Function

    End If

Next x

Set myTable = Nothing

'MsgBox ("ERR.GetColNumberFromName: no matching found")

End Function

' ==========================================================
Public Function GetValueFromTable_offset(myTable As ListObject, colNum As Variant, varName As String, iOffset As Long) As Variant
'
Dim myArray As Variant
Dim x As Long
'
'Create Array List from Table
myArray = myTable.DataBodyRange
'Looping for matching
For x = LBound(myArray) To UBound(myArray)
        If myArray(x, 1 + iOffset) = varName Then
            GetValueFromTable_offset = myTable.DataBodyRange(x, colNum).Value
            Exit Function
        End If
Next x
'
GetValueFromTable_offset = "ERR.GetValueFromTable_offset: no matching found"
'
End Function
Sub RemoveTableBodyData(tbl As ListObject)
'Delete Table's Body Data
  If tbl.ListRows.Count >= 1 Then
    tbl.DataBodyRange.Delete
  End If
'
End Sub
Public Function DetermineActiveTable(TableName As String) As Boolean
'  Determine if ActiveCell is inside a Table
'            = TRUE, success
'            = FALSE, fail
'
Dim SelectedCell As Range
Dim ActiveTable As ListObject
'
Set SelectedCell = ActiveCell


    On Error GoTo NoTableSelected

    TableName = SelectedCell.ListObject.Name
    Set ActiveTable = ActiveSheet.ListObjects(TableName)
    On Error GoTo 0
'
    DetermineActiveTable = True
    Set SelectedCell = Nothing
  
Exit Function
'Error Handling
NoTableSelected:
    DetermineActiveTable = False
    Set SelectedCell = Nothing
    
End Function
Sub CopyPasteRange()
Dim NbCol As Integer
Dim NbRow As Integer
Dim CopyRange As Range 'Plage de cellules que l'on veux copier
Dim PasteRange As Range 'Plage de cellules ou mettre les informations
 
'ThisWorkbook = Comme son nom l'indique c'est le classeur à partir duquel le code est lancé,
'si on en veux un autre il faut utiliser le nom du classeur: excel.Application.Workbooks("monfichierExcel.xls")
'
'Sheets(1) = Feuil N°1 du classeur selectionné, on peut remplacé par le nom de la feuil exemple: "Récapitulation"
'
'Cells(ligne,colonne) = la cellule de départ.
'
'Resize(ligne,colonne) = selectionne de la cellule de debut jusqu'a la cellule specifié par la ligne et la colonne.
'
'ThisWorkbook.Sheets(1).Cells(iCompteur2, 1).Resize(iCompteur2, Colmax) = Selectionne une plage de cellules
'
'Le veux dire: je sauvegarde la reference de cette plage de cellule dans CopyRange
Set CopyRange = ThisWorkbook.Sheets(1).Cells(iCompteur2, 1).Resize(iCompteur2, Colmax)
 
'Compte le nombre de lignes dans la reference de la plage de cellule de copyrange
NbRow = myRange.Rows.Count
 
'Compte le nombre de colonnes dans la reference de la plage de cellule de copyrange
NbCol = myRange.Columns.Count
 
'With = est une sorte de raccourcie, au lieu d'ecrire a chaque fois [ThisWorkbook.Sheets("Récapitulation")] on l'écrit une fois
'puis lorsque l'on met un [.] ca veux dire que l'on réutilise ce racourcie, ca permet d'avoir un code plus claire.
'Explication:
'With ThisWorkbook.Sheets("Récapitulation")
'Set PasteRange = .Range(.Cells(iIndexTab, 1), .Cells(NbCol, NbRow))
'End With
'Veux dire:
'Set PasteRange = ThisWorkbook.Sheets("Récapitulation").Range(ThisWorkbook.Sheets("Récapitulation").Cells(iIndexTab, 1), ThisWorkbook.Sheets("Récapitulation").Cells(NbCol, NbRow))
'
'Ce code a pour but de sauvegarder la reference d'une plage de cellule égale à celle de CopyRange dans la feuil désiré.
With ThisWorkbook.Sheets("Récapitulation")
Set PasteRange = .Range(.Cells(iIndexTab, 1), .Cells(NbCol, NbRow))
End With
End Sub
Function uniqueCode(fileName As String, fileDate As Date, fileSize As Long) As String
'
Dim i As Long, iSum As Long
For i = 1 To Len(fileName)
    iSum = iSum + Asc(Mid$(fileName, i, 1)) * i
Next i
'
    uniqueCode = Hex(iSum) & "-" & _
                 Hex((Year(fileDate) - 1900)) & _
                 Hex(Month(fileDate)) & _
                 Hex(Day(fileDate)) & _
                 Hex(Hour(fileDate)) & _
                 Hex(Minute(fileDate)) & _
                 Hex(Second(fileDate)) & "-" & _
                 Hex(fileSize)
    
End Function
Function FileNameNoExt(strPath As String) As String
    Dim strTemp As String
    strTemp = Mid$(strPath, InStrRev(strPath, "\") + 1)
    FileNameNoExt = Left$(strTemp, InStrRev(strTemp, ".") - 1)
End Function
 
 'The following function returns the filename with the extension from the file's full path:
Function FileNameWithExt(strPath As Variant) As String
    FileNameWithExt = Mid$(strPath, InStrRev(strPath, "\") + 1)
End Function
 
 'the following function will get the path only (i.e. the folder) from the file's ful path:
Function FilePath(strPath As String) As String
    FilePath = Left$(strPath, InStrRev(strPath, "\"))
End Function
Function randLong(ByVal maxRandom As Long) As Long
'
Randomize
randLong = Int(maxRandom * Rnd)
'
End Function
Function copySheetOnNewFile(typeSend As String, sheetName As String, fPath As String, fName As String) As Boolean
'   copy sheet (sheetName) on a file (fpath+fName) and transform their cells in values
'   typeSend,
'   sheetName,
'   nameNewSheet,
'   fPath,
'   fName,
'   myPassword,
'
    Dim newBook As Workbook, allSheet As Range
'
    Set newBook = Workbooks.Add
'
    ThisWorkbook.Sheets(sheetName).Copy Before:=newBook.Sheets(1)
'
    If Dir(fPath & "\" & fName) <> "" Then
        MsgBox "Fichier " & fPath & "\" & fName & " existe déjà"
        copySheetOnNewFile = False
        Set newBook = Nothing
'
    Else
        newBook.SaveAs fileName:=fPath & "\" & fName
'       select all cells of copied sheet
        Set allSheet = Worksheets(sheetName).Cells
'        allSheet.Select
'       copy all cells and past as values
        allSheet.Copy
        allSheet.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                               SkipBlanks:=False, Transpose:=False
        copySheetOnNewFile = True
        Set newBook = Nothing
        Set allSheet = Nothing
'
    End If
'
End Function
Public Function whoIam() As String
'
Dim myTable As ListObject
Dim x As Long
'Set path for Table variable
Set myTable = Sheets("Data").ListObjects("T_noms")
'
    Application.Volatile
'    Range("N8").Value = ActiveCell.Address
    If (DetermineActiveTable("T_horsBAG")) Then
        Select Case Mid(ActiveCell.Address, 1, 3)
        Case "$M$", "$C$"
            whoIam = GetValueFromTable_offset(myTable, 2, ActiveCell.Cells, 3) & " (" & GetValueFromTable_offset(myTable, 5, ActiveCell.Cells, 3) & ")"
        
        Case Else
            whoIam = ""
        
        End Select
    
    Else
        whoIam = ""
    
    End If
 '
 Set myTable = Nothing
    
End Function
Public Function Progress(status1 As String, status2 As String) As Single
'
Dim myTable As ListObject
Dim x As Long, addEXE As Single
'
    addEXE = 0#
    If UCase(Trim(status1)) = "ANNULE" Then addEXE = 0.25
    Progress = GetValueFromTable_Data("T_état", 2, UCase(Trim(status1))) * 0.25 + GetValueFromTable_Data("T_état", 2, UCase(Trim(status2))) * (0.75 + addEXE)
'
End Function
Public Function quadrantRLA(successFail As String, noteSF As String, plannedNonPlanned As String) As String
Dim index As Integer, adverb As String
'
     index = 0
'
    Select Case successFail
    Case "succès"
        index = 1000
        
    Case "échec"
        index = 100
        
    End Select
 '
    Select Case plannedNonPlanned
    Case "planifié"
        index = index + 10
    
    Case "non-planifié"
        index = index + 1
    
    End Select
'
    Select Case noteSF
        Case "passable"
            adverb = " (SURELY)"
        
        Case "faible"
            adverb = " (JUSTLY)"
        
        Case "assez bien"
            adverb = " (VERY)"
        
        Case "moindre"
            adverb = " (SLIGHTLY)"
            
        Case "bien"
            adverb = " (EASILY)"
        
        Case "grave"
            adverb = " (MUCH)"
        
        Case "très bien"
            adverb = " (JUSTLY)"
            
        Case "très grave"
            adverb = " (ABSOLUTELY)"
        
        Case Else
            adverb = ""
            
    End Select
'
    If index = 0 Then
        quadrantRLA = "NOT RATED" & adverb
    
    ElseIf index = 1010 Then
        quadrantRLA = "KEEP" & adverb
    
    ElseIf index = 1001 Then
        quadrantRLA = "ADD" & adverb
        
    ElseIf index = 110 Then
        quadrantRLA = "IMPROVE" & adverb
        
    ElseIf index = 101 Then
        quadrantRLA = "DROP" & adverb
    
    Else
        quadrantRLA = "---"
     
    End If
    
End Function
Public Function scoreRLA(successFail As String, noteSF As String, lot As String) As Single
Dim index As Integer, adverb As String, adval As Single
'
     scoreRLA = 0
'
    Select Case lot
    Case "BAG"
        addVal = 0.3
        
    Case "ALL"
        addVal = 0
        
    Case "BAT"
        addVal = -0.3
    
    End Select
'
    Select Case successFail
    Case "succès"
        Select Case noteSF
            Case "passable"
                scoreRLA = 1
              
            Case "assez bien"
                scoreRLA = 2
        
            Case "bien"
                scoreRLA = 3
        
            Case "très bien"
                scoreRLA = 4
            
            End Select
        
    Case "échec"
        Select Case noteSF
            Case "faible"
                scoreRLA = -1
              
            Case "moindre"
                scoreRLA = -2
        
            Case "grave"
                scoreRLA = -3
        
            Case "très grave"
                scoreRLA = -4
            
            End Select
        
    End Select
    
    scoreRLA = scoreRLA + addVal
    
End Function
Sub EnvoyerEmail(ByVal Sujet As String, ByVal Destinataire As String, ByVal ContenuEmail As String, Optional ByVal PieceJointe As String)
'par Excel-Malin.com ( https://excel-malin.com )

On Error GoTo EnvoyerEmailErreur

'définition des variables
Dim oOutlook As Outlook.Application
Dim WasOutlookOpen As Boolean
Dim oMailItem As Outlook.MailItem
Dim Body As Variant

Body = ContenuEmail

    'vérification si le Contenu du mail n'est pas vide. Si oui, email n'est pas envoyé. Si vous voulez pouvoir envoyer les email vides, mettez en commentaire les 4 lignes de code qui suivent.
    If (Body = False) Then
        MsgBox "Mail non envoyé car vide", vbOKOnly, "Message"
        Exit Sub
       End If
    
    'préparer Outlook
    PreparerOutlook oOutlook
    Set oMailItem = oOutlook.CreateItem(0)
    
    'création de l'email
    With oMailItem
        .To = Destinataire
        .Subject = Sujet
        
        'CHOIX DU FORMAT
        '----------------------
        'email formaté comme texte
            .BodyFormat = olFormatRichText
            .Body = Body
            
            'OU
            
        'email formaté comme HTML
            '.BodyFormat = olFormatHTML
            '.HTMLBody = "<html><p>" & Body & "</p></html>"
        '----------------------
        
        If PieceJointe <> "" Then .Attachments.Add PieceJointe

       .Display   '<- affiche l'email (si vous ne voulez pas l'afficher, mettez cette ligne en commentaire)
       .Save      '<- sauvegarde l'email avant l'envoi (pour ne pas le sauvegarder, mettez cette ligne en commentaire)
       .Send      '<- envoie l'email (si vous voulez seulement préparer l'email et l'envoyer manuellement, mettez cette ligne en commentaire)
    End With
    
   'nettoyage...
    If (Not (oMailItem Is Nothing)) Then Set oMailItem = Nothing
    If (Not (oOutlook Is Nothing)) Then Set oOutlook = Nothing
    
   Exit Sub

EnvoyerEmailErreur:
    If (Not (oMailItem Is Nothing)) Then Set oMailItem = Nothing
    If (Not (oOutlook Is Nothing)) Then Set oOutlook = Nothing
  
    MsgBox "Le mail n'a pas pu être envoyé...", vbCritical, "Erreur"
End Sub
Private Sub PreparerOutlook(ByRef oOutlook As Object)
'par Excel-Malin.com ( https://excel-malin.com )

'------------------------------------------------------------------------------------------------
'Ce code vérifie si Outlook est prêt à envoyer des emails... Et s'il ne l'est pas, il le prépare.
'------------------------------------------------------------------------------------------------
On Error GoTo PreparerOutlookErreur


On Error Resume Next
    'vérification si Outlook est ouvert
    Set oOutlook = GetObject(, "Outlook.Application")
    
    If (Err.Number <> 0) Then 'si Outlook n'est pas ouvert, une instance est ouverte
        Err.Clear
        Set oOutlook = CreateObject("Outlook.Application")
    Else    'si Outlook est ouvert, l'instance existante est utilisée
        Set oOutlook = GetObject("Outlook.Application")
        oOutlook.Visible = True
    End If
    Exit Sub

PreparerOutlookErreur:
    MsgBox "Une erreur est survenue lors de l'exécution de PreparerOutlook()..."
End Sub
Public Sub insertComment(rg As Range, Optional txt As String = "")
'------------------------------------
'Insertion of notes
'------------------------------------

  With rg
    If .comment Is Nothing Then
        If txt <> "" Then
            .AddComment
            .comment.Text txt
        
        End If
    
    Else
        If txt = "" Then
            .comment.Delete
        
        Else
            .comment.Text txt
        End If
       
    
    End If

  End With

End Sub

