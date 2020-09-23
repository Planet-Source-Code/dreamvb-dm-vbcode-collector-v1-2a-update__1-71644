VERSION 5.00
Begin VB.Form frmMove 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move Code"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.dFlatButton cmdCancel 
      Height          =   345
      Left            =   2430
      TabIndex        =   3
      Top             =   795
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
   Begin Project1.dFlatButton cmdOk 
      Height          =   350
      Left            =   1260
      TabIndex        =   2
      Top             =   795
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
      Caption         =   "OK"
   End
   Begin VB.ComboBox cboMove 
      Height          =   315
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3285
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move item To:"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   135
      Width           =   1020
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GetCatNames(TCombo As ComboBox)
Dim rc As Recordset
    
    'Load record set
    Set rc = db.OpenRecordset("Category")
    
    While Not rc.EOF
        If Not LCase(mCatName) = LCase(rc.Fields("CatName").Value) Then
            'Add categorys
            TCombo.AddItem rc.Fields("CatName").Value
        End If
        rc.MoveNext
    Wend
    
    If (TCombo.ListCount) Then
        TCombo.ListIndex = 0
    End If
    
End Sub

Private Sub cboMove_Click()
    'Set category
    mCatName = cboMove.Text
End Sub

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmMove
End Sub

Private Sub cmdok_Click()
    ButtonPress = vbOK
    Unload frmMove
End Sub

Private Sub Form_Load()
    Set frmMove.Icon = Nothing
    'Load category list
    Call GetCatNames(cboMove)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMove = Nothing
End Sub
