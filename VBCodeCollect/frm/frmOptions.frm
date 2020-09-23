VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkTray 
      Caption         =   "Always minsize to tray"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   915
      Width           =   4350
   End
   Begin Project1.dFlatButton cmdok 
      Height          =   345
      Left            =   3600
      TabIndex        =   4
      Top             =   1395
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
   Begin Project1.dFlatButton cmdCancel 
      Height          =   345
      Left            =   4785
      TabIndex        =   3
      Top             =   1395
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
   Begin Project1.dFlatButton cmdOpen 
      Height          =   390
      Left            =   5235
      TabIndex        =   2
      ToolTipText     =   "Open"
      Top             =   405
      Width           =   570
      _ExtentX        =   1005
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
      Caption         =   ". . ."
   End
   Begin VB.TextBox txtPath 
      Height          =   360
      Left            =   270
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   4890
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   255
      X2              =   5820
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   255
      X2              =   5820
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Always load of following database on start-up:"
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3225
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload frmOptions
End Sub

Private Sub cmdok_Click()
    'Add database path to regedit
    SaveSetting "DMCodeBank", "cfg", "dbPath", txtPath.Text
    SaveSetting "DMCodeBank", "cfg", "Minsize", chkTray.Value
    Call cmdCancel_Click
End Sub

Private Sub cmdOpen_Click()
Dim lFile As String
    'Get Filename
    lFile = frmmain.GetOpenDLGName(Filter1)
    
    If Len(lFile) Then
        txtPath.Text = lFile
    End If
End Sub

Private Sub Form_Load()
    Set frmOptions.Icon = Nothing
    'Update text box with database path
    txtPath.Text = GetSetting("DMCodeBank", "cfg", "dbPath", FixPath(App.Path) & "codetips.mdb")
    chkTray.Value = GetSetting("DMCodeBank", "cfg", "Minsize", 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOptions = Nothing
End Sub
