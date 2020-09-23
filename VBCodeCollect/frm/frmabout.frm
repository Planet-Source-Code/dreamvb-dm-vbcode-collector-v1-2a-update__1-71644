VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.dFlatButton cmdok 
      Height          =   345
      Left            =   3345
      TabIndex        =   4
      Top             =   1710
      Width           =   915
      _ExtentX        =   1614
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM Source code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   735
      TabIndex        =   5
      Top             =   210
      Width           =   1770
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   60
      X2              =   4305
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created By DreamVB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   60
      TabIndex        =   3
      Top             =   1860
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "THIS PROGRAM IS FREEWARE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   795
      TabIndex        =   2
      Top             =   1095
      Width           =   2775
   End
   Begin VB.Image ImgLogo 
      Height          =   465
      Left            =   105
      Top             =   135
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The perfect solution for organising your important source codes"
      Height          =   495
      Left            =   825
      TabIndex        =   1
      Top             =   585
      Width           =   3285
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Collector V1.2a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A56E3A&
      Height          =   270
      Left            =   2535
      TabIndex        =   0
      Top             =   225
      Width           =   1590
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdok_Click()
    Unload frmabout
End Sub

Private Sub Form_Load()
    Set frmabout.Icon = Nothing
    Set ImgLogo.Picture = frmmain.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing
End Sub
