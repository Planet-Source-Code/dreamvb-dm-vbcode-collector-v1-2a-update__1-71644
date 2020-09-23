VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dZipView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const SIG = &H4034B50
Private Const ENDSIG = &H2014B50

Private Type ZFHeader
    Signature As Long
    Version As Integer
    GPBFlag As Integer
    Compress As Integer
    DateTime As Long
    CRC32 As Long
    CSize As Long
    USize As Long
    FileNameLen As Long
    ExtraField As Integer
End Type

Public Function GetOpenZipFiles(ByVal Filename As String) As Collection
Dim fp As Long
Dim fName As String
Dim zInfo As ZFHeader
Dim Offset As Long
Dim Col As New Collection
Dim sPos As Integer

    fp = FreeFile
    
    Open Filename For Binary As #fp
        Do
            'Get header
            Get #fp, , zInfo
            'Exit if end signature found
            If (zInfo.Signature = ENDSIG) Then Exit Do
            'Check for correct zip signature
            If (zInfo.Signature = SIG) Then
                fName = Space(zInfo.FileNameLen)
                Offset = Seek(fp) - 2
                'Get Filename
                Get fp, Offset, fName
                sPos = InStr(1, fName, "/", vbBinaryCompare)
                If (sPos > 1) Then
                    'Add filename to collection
                    Col.Add Trim(Mid(fName, sPos + 1))
                End If
            End If
            'Inc offset
            Offset = (Seek(fp) - 1) + zInfo.CSize + 1
            'move to next offsert
            Seek #fp, Offset
        Loop Until (zInfo.Signature <> SIG)
    Close #fp
    
    Set GetOpenZipFiles = Col
    Set Col = Nothing
    
End Function
