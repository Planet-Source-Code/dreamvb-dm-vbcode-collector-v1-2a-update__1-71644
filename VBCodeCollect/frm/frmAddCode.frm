VERSION 5.00
Begin VB.Form frmAddCode 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.dFlatButton cmdCancel 
      Height          =   350
      Left            =   6150
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5805
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
   End
   Begin Project1.dFlatButton cmdClear 
      Height          =   350
      Left            =   5025
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5805
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clear"
   End
   Begin Project1.dFlatButton cmdAdd 
      Height          =   350
      Left            =   3900
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5805
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "#0"
   End
   Begin Project1.dEditor txtCode 
      Height          =   2640
      Left            =   60
      TabIndex        =   14
      Top             =   3075
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   4657
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtVersion 
      Height          =   330
      Left            =   3480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1215
      Width           =   1680
   End
   Begin Project1.dFlatButton cmdBut 
      Height          =   390
      Index           =   3
      Left            =   1320
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Paste"
      Top             =   2640
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmAddCode.frx":0000
   End
   Begin Project1.dFlatButton cmdBut 
      Height          =   390
      Index           =   0
      Left            =   60
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Open"
      Top             =   2640
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmAddCode.frx":0352
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   75
      TabIndex        =   7
      Top             =   0
      Width           =   7170
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "#1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   795
         TabIndex        =   8
         Top             =   315
         Width           =   360
      End
      Begin VB.Image imgEdit 
         Height          =   480
         Left            =   195
         Picture         =   "frmAddCode.frx":06A4
         Top             =   225
         Width           =   480
      End
   End
   Begin VB.TextBox txtAuthor 
      Height          =   330
      Left            =   5235
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1215
      Width           =   1965
   End
   Begin VB.TextBox txtcomment 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1830
      Width           =   7155
   End
   Begin VB.TextBox txtTitle 
      Height          =   330
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1215
      Width           =   3360
   End
   Begin Project1.dFlatButton cmdBut 
      Height          =   390
      Index           =   1
      Left            =   480
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Cut"
      Top             =   2640
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmAddCode.frx":0C0B
   End
   Begin Project1.dFlatButton cmdBut 
      Height          =   390
      Index           =   2
      Left            =   900
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Copy"
      Top             =   2640
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmAddCode.frx":0F5D
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      Height          =   195
      Index           =   4
      Left            =   5220
      TabIndex        =   5
      Top             =   990
      Width           =   510
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   1605
      Width           =   705
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   2
      Top             =   990
      Width           =   570
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Title:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   990
      Width           =   765
   End
End
Attribute VB_Name = "frmAddCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SelectAllText(TBox As TextBox)
    'Select all text using ctrl+a
    TBox.SelStart = 0
    TBox.SelLength = Len(TBox.Text)
    TBox.SetFocus
End Sub

Private Sub LoadSniplet(ByVal Filename)
Dim fp As Long
    fp = FreeFile
    
    Open Filename For Binary As #fp
        Get #fp, , m_snip
    Close #fp
    'Fill in the text fieds
    
    With m_snip
        txtTitle.Text = RTrim(.dTitle)
        txtVersion.Text = RTrim(.dVersion)
        txtAuthor.Text = RTrim(.dAuthor)
        txtcomment.Text = .dComment
        txtCode.Text = .dCode
    End With
    
    'Clear up
    With m_snip
        .dAuthor = vbNullString
        .dCode = vbNullString
        .dComment = vbNullString
        .dVersion = vbNullString
        .dTitle = vbNullString
    End With
End Sub

Private Sub cmdAdd_Click()
    ButtonPress = vbOK
    'Code type
    With m_tcode
        .dTitle = txtTitle.Text
        .dVersion = txtVersion.Text
        .dCodeBlock = txtCode.Text
        .dComment = txtcomment.Text
        .dAuthor = txtAuthor.Text
        
        If Len(m_tcode.dAuthor) = 0 Then
            .dAuthor = "None"
        End If
    
        If Len(.dComment) = 0 Then
            .dComment = "None"
        End If
        
        If Len(.dVersion) = 0 Then
            .dVersion = "None"
        End If
        
    End With
    
    'Unload the form
    Unload frmAddCode
End Sub

Private Sub cmdBut_Click(Index As Integer)
Dim lFile As String

    Select Case Index
        Case 0 'Open
            'Get Filename
            lFile = frmmain.GetOpenDLGName(Filter2)

            If Len(lFile) Then
                If (frmmain.mFilterIdx = 2) Then
                    'Load snipplet file
                    Call LoadSniplet(lFile)
                Else
                    txtCode.Text = OpenFile(lFile)
                End If
            End If
        Case 1 'Cut
            Clipboard.SetText txtCode.SelText
            txtCode.SelText = ""
        Case 2 'Copy
            Clipboard.SetText txtCode.SelText
        Case 3 'Paste
            txtCode.SelText = Clipboard.GetText(vbCFText)
            txtCode.SetFocus
    End Select
    
    txtCode.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    'Unload the form
    Unload frmAddCode
End Sub

Private Sub cmdClear_Click()
    txtTitle.Text = ""
    txtAuthor.Text = ""
    txtcomment.Text = ""
    txtVersion.Text = ""
    txtCode.Text = ""
    txtTitle.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    txtTitle.SetFocus
End Sub

Private Sub Form_Load()

    Set frmAddCode.Icon = Nothing
    
    If (EditOp = 1) Then
        frmAddCode.Caption = "Edit"
        cmdAdd.Caption = "Update"
        lblTitle(5).Caption = "Modify your sourcecode"
        'Update controls
        With m_tcode
            txtTitle.Text = .dTitle
            txtCode.Text = .dCodeBlock
            txtcomment.Text = .dComment
            txtAuthor.Text = .dAuthor
            txtVersion.Text = .dVersion
        End With
    Else
        frmAddCode.Caption = "New"
        cmdAdd.Caption = "Add"
        lblTitle(5).Caption = "Add new sourcecode"
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAddCode = Nothing
End Sub

Private Sub txtAuthor_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 1) Then
        Call SelectAllText(txtAuthor)
        KeyAscii = 0
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 1) Then
        Call SelectAllText(txtCode)
        KeyAscii = 0
    End If
End Sub

Private Sub txtcomment_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 1) Then
        Call SelectAllText(txtcomment)
        KeyAscii = 0
    End If
End Sub

Private Sub txtTitle_Change()
    cmdAdd.Enabled = Len(Trim$(txtTitle.Text)) > 0
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 1) Then
        Call SelectAllText(txtTitle)
        KeyAscii = 0
    End If
End Sub

Private Sub txtVersion_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 1) Then
        Call SelectAllText(txtVersion)
        KeyAscii = 0
    End If
End Sub
