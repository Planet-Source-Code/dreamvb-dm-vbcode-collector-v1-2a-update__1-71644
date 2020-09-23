Attribute VB_Name = "DbHelper"
Option Explicit
Public db As Database

Public Function CreateNewDatabase(ByVal Filename As String) As Integer
Dim tmp As Database
Dim td As TableDef
On Error GoTo CreateErr:
    'Create the new blank database
    Set tmp = CreateDatabase(Filename, dbLangGeneral, dbVersion30)
    'open Database
    Set tmp = OpenDatabase(Filename, False)
    
    'Create Table
    Set td = tmp.CreateTableDef("Category")
    
    With td
        'Add field info
        .Fields.Append .CreateField("ID", dbLong, 4)
        'Below is needed for autonumbers dblong+attr
        .Fields(0).Attributes = 49
        .Fields.Append .CreateField("CatName", dbText, 50)
        .Fields.Append .CreateField("Parent", dbInteger, 2)
    End With
    
    'Append the field info to the table
    Call tmp.TableDefs.Append(td)
    
    'Add Table
    Set td = tmp.CreateTableDef("Codes")
    
    With td
        'Add field's
        .Fields.Append .CreateField("ID", dbLong, 4)
        'Below is needed for autonumbers dblong+attr
        .Fields(0).Attributes = 49
        .Fields.Append .CreateField("Title", dbText, 50)
        .Fields.Append .CreateField("Version", dbText, 50)
        .Fields.Append .CreateField("CatID", dbInteger, 2)
        .Fields.Append .CreateField("sCode", dbMemo)
        'Allow code field to have zero length
        .Fields(4).AllowZeroLength = True
        .Fields.Append .CreateField("sComment", dbMemo)
        .Fields.Append .CreateField("sAuthor", dbText, 50)
    End With
    
    'Append table to the database
    Call tmp.TableDefs.Append(td)
    CreateNewDatabase = 1
    'Clear up
    tmp.Close
    Set tmp = Nothing
    Set td = Nothing
    Exit Function
CreateErr:
    CreateNewDatabase = 0
End Function

Public Function dOpenDataBase(ByVal Filename As String) As Boolean
On Error GoTo OpenDBErr:
    'Check if this database is set to readonly
    
    Set db = OpenDatabase(Filename, False, isReadOnly(Filename))
    'Tell us the database is open
    dOpenDataBase = True
    Exit Function
OpenDBErr:
    dOpenDataBase = False
End Function

Public Sub LoadCategories(TTreeView As TreeView)
Dim rc As Recordset
Dim CatID As Integer
Dim pId As String
Dim nNode As Node
On Error Resume Next

    'Open the Category table
    Set rc = db.OpenRecordset("Category")

    'Clear nodes
    With TTreeView.Nodes
        .Clear
        'Add first node
        .Add , tvwFirst, "M_TOP", "Code Collector", 6, 6
        'Loop tho the records
        
        While (Not rc.EOF)
            pId = rc.Fields("Parent").Value
            CatID = rc.Fields("ID").Value
            'Create Node to add
            Set nNode = .Add("M_TOP", tvwChild, "C" & CatID, rc.Fields("CatName").Value, 2, 2)
            nNode.Tag = "C" & pId
            'Get next record
            rc.MoveNext
        Wend
    End With

    For Each nNode In TTreeView.Nodes
        pId = nNode.Tag
        If Len(pId) <> 0 Then
            If pId = "C0" Then pId = "M_TOP"
            Set nNode.Parent = TTreeView.Nodes(pId)
        End If
    Next
    
    'Expand only the first node and close the rest
    For Each nNode In TTreeView.Nodes
        If (nNode.Index <> 1) Then
            nNode.Expanded = Not nNode.Expanded
            nNode.Tag = vbNullString
        End If
    Next
    
    'Clear up
    Set nNode = Nothing
    Set rc = Nothing
    pId = vbNullString
    
End Sub

Public Sub LoadRecords(TListView As ListView, ByVal pId As Long)
Dim rc As Recordset
Dim sQuery As String

    'Build the Query string
    sQuery = "SELECT ID,Title,Version,CatID,sAuthor FROM Codes WHERE CatID = " & Str(pId)
    
    With TListView
        'Clear the control's data
        .ListItems.Clear
        .Sorted = False
        'Create the record set
        Set rc = db.OpenRecordset(sQuery)
        If (rc.RecordCount) Then
            'Loop tho all the records
            While (Not rc.EOF)
                'add first item
                .ListItems.Add , "k," & rc.Fields("ID").Value, rc.Fields("Title").Value, 9, 9
                'Add remaining subitems
                .ListItems(.ListItems.Count).SubItems(1) = rc.Fields("Version").Value
                .ListItems(.ListItems.Count).SubItems(2) = rc.Fields("sAuthor").Value & ""
                'Get next record
                rc.MoveNext
            Wend
        End If
        .Refresh
        .Sorted = True
        
    End With
    
    'Clean up
    rc.Close
    Set rc = Nothing
    sQuery = vbNullString
End Sub

Public Function AddCategory(ByVal CategoryName As String, Optional ByVal ParentID As Integer) As Integer
Dim rc As Recordset
On Error GoTo AddErr:
    'Open the record set
    Set rc = db.OpenRecordset("Category")
    
    With rc
        'Add new record
        .AddNew
        !CatName = CategoryName
        !Parent = ParentID
        .Update
    End With
    
    'Record was added
    Set rc = Nothing
    AddCategory = 1
    Exit Function
    'Error flag
AddErr:
    AddCategory = 0
End Function

Public Function EditCategory(ByVal ID As Long, ByVal NewName As String) As Integer
Dim rc As Recordset
On Error GoTo EditErr:
    'Open the record set
    Set rc = db.OpenRecordset("SELECT ID,CatName FROM Category WHERE ID = " & Str(ID))
    
    'Edit the category name
    With rc
        .Edit
        !CatName = NewName
        .Update
    End With
    
    Set rc = Nothing
    EditCategory = 1
    
    Exit Function
    'Error flag
EditErr:
    EditCategory = 0
End Function

Public Function DeleteCategory(nNode As Node) As Integer
Dim rc As Recordset
Dim rc2 As Recordset
Dim cNode As Node
Dim ID As Integer
On Error GoTo DelErr:

    'Get child node
    Set cNode = nNode.Child
    'Get ID
    ID = CInt(Mid(nNode.Key, 2))
    'Open the recordset
    Set rc = db.OpenRecordset("SELECT ID,CatName,Parent FROM Category WHERE ID = " & Str(ID))
    'Delete the the record
    rc.Delete
    'Open the codes table
    Set rc2 = db.OpenRecordset("codes")
    'Delete the records
    Do While (Not rc2.EOF)
        With rc2
            If (!CatID) = ID Then .Delete
            'Get next record
            .MoveNext
        End With
    Loop
    'Loop tho the child nodes
    Do While Not (cNode Is Nothing)
        'Preform the delete
        Call DeleteCategory(cNode)
        Set cNode = cNode.Next
    Loop
    
    'Clear up
    Set rc = Nothing
    Set rc2 = Nothing
    Set cNode = Nothing
    
    DeleteCategory = 1
    Exit Function
    'Error flag
DelErr:
    DeleteCategory = 0
End Function

Public Sub DeleteRecords(ByVal pId As Long)
Dim sQuery As String
On Error Resume Next
    'Delete all records from codes database based on the CatID
    sQuery = "DELETE ID,Title,Version,CatID,sCode,sComment,sAuthor FROM codes WHERE CatID = " & Str(pId)
    'execute the command
    db.Execute sQuery
End Sub

Public Function AddNewCode(ByVal pId As Long) As Long
Dim rc As Recordset
On Error GoTo CodeAddErr:
    'Open the codes record set
    Set rc = db.OpenRecordset("codes")
    
    With rc
        'Add the new record
        .AddNew
        !Title = m_tcode.dTitle
        !Version = m_tcode.dVersion
        !CatID = pId
        !scode = m_tcode.dCodeBlock
        !sComment = m_tcode.dComment
        !sAuthor = m_tcode.dAuthor
        .Update
    End With
    
    AddNewCode = 1
    'Clear up
    m_tcode.dCodeBlock = vbNullString
    m_tcode.dTitle = vbNullString
    m_tcode.dVersion = vbNullString
    
    Set rc = Nothing
    Exit Function
CodeAddErr:
    AddNewCode = 0
End Function

Public Function DeleteCode(ByVal ID As Long) As Long
Dim sQuery As String
Dim rc As Recordset
On Error GoTo DelCodeErr:

    sQuery = "SELECT ID,Title,Version,CatID,sCode,sComment,sAuthor FROM codes WHERE ID = " & Str(ID)
    'Open record set
    Set rc = db.OpenRecordset(sQuery)
    'Check for record count
    If (rc.RecordCount) Then
        'Delete the record
        rc.Delete
    End If
    
    DeleteCode = 1
    Set rc = Nothing
    
    Exit Function
DelCodeErr:
    DeleteCode = 0
End Function

Public Function EditCode(ByVal ID As Long) As Long
Dim sQuery As String
Dim rc As Recordset
On Error GoTo EditCodeErr:

    sQuery = "SELECT ID,Title,Version,CatID,sCode,sComment,sAuthor FROM codes WHERE ID = " & Str(ID)
    'Open record set
    Set rc = db.OpenRecordset(sQuery)
    'Edit the code
    With rc
        .Edit
        !Title = m_tcode.dTitle
        !Version = m_tcode.dVersion
        !scode = m_tcode.dCodeBlock
        !sComment = m_tcode.dComment
        !sAuthor = m_tcode.dAuthor
        .Update
    End With
    
    EditCode = 1
    Set rc = Nothing
    
    Exit Function
EditCodeErr:
    EditCode = 0
End Function

Public Function RecordCount(ByVal TableName As String) As Long
Dim rc As Recordset
On Error GoTo OpenErr:
    'Open the recordset
    Set rc = db.OpenRecordset(TableName)
    'Return record count
    RecordCount = rc.RecordCount
    
    Set rc = Nothing
    Exit Function
    'Error flag
OpenErr:
    RecordCount = -1
End Function

Public Function MoveCodeItem(ByVal ToLoc As String, ID As Long) As Integer
Dim rc As Recordset
Dim sQuery As String
Dim pId As Long
Dim Ret As Integer
    
    sQuery = "SELECT ID,CatName FROM Category WHERE CatName = '" & ToLoc & "'"
    'Open recordset
    Set rc = db.OpenRecordset(sQuery)
    'Get Code ParentID
    pId = CLng(rc.Fields("ID").Value)
    
    sQuery = "SELECT ID,Title,Version,CatID,sCode,sComment,sAuthor FROM codes WHERE ID = " & Str(ID)
    'Open the record set
    Set rc = db.OpenRecordset(sQuery)
    'Move the record over
    With m_tcode
        .dAuthor = rc.Fields("sAuthor").Value
        .dComment = rc.Fields("sComment").Value
        .dCodeBlock = rc.Fields("sCode").Value
        .dVersion = rc.Fields("Version").Value
        .dTitle = rc.Fields("Title")
    End With
    
    'Add the record
    If AddNewCode(pId) <> 1 Then
        Ret = 0
    ElseIf DeleteCode(ID) <> 1 Then
        Ret = 0
    Else
        Ret = 1
    End If
    
    MoveCodeItem = Ret
    'Clear up
    Set rc = Nothing
    sQuery = vbNullString
    
End Function
