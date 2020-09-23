Attribute VB_Name = "Tools"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long

Private Type TCode
    dTitle As String
    dVersion As String
    dCodeBlock As String
    dComment As String
    dAuthor As String
End Type

Private Type TSnipplet
    dTitle As String * 30
    dVersion As String * 30
    dAuthor As String * 30
    dComment As String
    dCode As String
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public m_tcode As TCode
Public m_snip As TSnipplet
Public ButtonPress As VbMsgBoxResult
Public EditOp As Integer '0=add,1=edit
Public m_SrcBrush As Long
Public mCatName As String

'Listview Consts
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Public Const Filter1 As String = "MDB Files(*.mdb)|*.mdb|"
Public Const Filter2 As String = "Text Files(*.txt)|*.txt|Snip File(*.snip)|*.snip|All Files(*.*)|*.*|"

Public Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Public Function FixPath(ByVal lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Function SaveText(ByVal Filename As String, fData As String)
Dim fp As Long
    fp = FreeFile
    'Saves text to a given filename.
    Open Filename For Output As #fp
        Print #fp, fData
    Close #fp
End Function

Public Function OpenFile(ByVal Filename As String) As String
Dim mBytes() As Byte
Dim fp As Long
    'Opens a file and returns it's data.
    fp = FreeFile
    Open Filename For Binary As #fp
        If LOF(fp) <> 0 Then
            ReDim Preserve mBytes(LOF(fp) - 1)
        End If
        'Get File bytes
        Get #fp, , mBytes
    Close #fp
    
    OpenFile = StrConv(mBytes, vbUnicode)
    Erase mBytes
    
End Function

Function isReadOnly(ByVal Filename As String) As Boolean
    isReadOnly = (GetAttr(Filename) = vbArchive) + vbReadOnly
End Function

Public Sub lvSizeColumns(lv As ListView)
Dim Counter As Long
    'Resizes Listview Column Headers.
    For Counter = 0 To (lv.ColumnHeaders.Count - 1)
        Call SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, Counter, _
        ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next Counter
End Sub

Public Sub TexturePicBox(PicBox As PictureBox)
Dim rc As RECT
    With PicBox
        rc.Right = (.ScaleWidth)
        rc.Bottom = (.ScaleHeight)
        FillRect .hDC, rc, m_SrcBrush
        '
        PicBox.Line (0, .ScaleHeight - 8)-(.ScaleWidth - 1, .ScaleHeight - 8), vbWhite
        .Refresh
    End With
End Sub

