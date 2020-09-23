VERSION 5.00
Begin VB.UserControl dFlatButton 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   79
   ToolboxBitmap   =   "dFlatButton.ctx":0000
   Begin VB.PictureBox PicImg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   660
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "dFlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" _
        (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc _
        As Long, ByVal lParam As String, ByVal wParam As Long, _
        ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, _
        ByVal n4 As Long, ByVal un As Long) As Long

'Flat Buton Style.
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4

Private Const BDR_SUNKEN95 As Long = &HA
Private Const BDR_RAISED95 As Long = &H5

Private Const BF_RECT As Long = &HF
'Text Consts
Private Const DST_PREFIXTEXT = &H2
Private Const DSS_NORMAL = &H0
Private Const DSS_DISABLED = &H20

Enum TButtonStyleA
    Win95 = 0
    Flat = 1
    Frame = 2
End Enum

Enum TAlign
    aLeft = 0
    aCenter = 1
    aRight = 2
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private m_GotFocus As Boolean
Private m_CRect As RECT
Private m_CapAlign As TAlign
Private m_Button As MouseButtonConstants

Private m_Caption As String
Private m_ShowRect As Boolean
Private b_Style As TButtonStyleA
'Event Declarations:
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Function DrawTextA(DrawOnDC As Long, X As Long, Y As Long, _
       hStr As String, tEnabled As Boolean, Clr As Long) As Long
Dim OT As Long

    If DrawOnDC = 0 Then
        Exit Function
    End If
    
    ' Set new text color and save the old one
    OT = GetTextColor(DrawOnDC)
    SetTextColor DrawOnDC, Clr
    ' Draw the text
    DrawTextA = DrawStateText(DrawOnDC, 0&, 0&, hStr, Len(hStr), _
               X, Y, 0&, 0&, DST_PREFIXTEXT Or IIf(tEnabled = True, _
               DSS_NORMAL, DSS_DISABLED))
    'Restore old text color
    SetTextColor DrawOnDC, OT
End Function

Private Sub BSetFocus()
    'Check if we have focus, and insure that m_ShowRect is enabled
    If (m_GotFocus) And (m_ShowRect) Then
        SetRect m_CRect, 3, 3, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3
        DrawFocusRect UserControl.hDC, m_CRect
    End If
End Sub

Private Sub DrawButton(Optional bState As Boolean = False)
Dim xImgPos As Integer
Dim yImgPos As Integer
Dim HasPic As Boolean
Dim BStyleDown As Long
Dim BStyleUp As Long

Const ImgWidth As Integer = 16
    
    With UserControl
        .Cls
        
        SetRect m_CRect, 0, 0, .ScaleWidth, .ScaleHeight
        
        'Image Positions
        xImgPos = 5
        yImgPos = (m_CRect.Bottom - ImgWidth) \ 2
        'Find out if we have a picture loaded.
        HasPic = (PicImg.Picture <> 0)
        
        Select Case b_Style
            Case Flat
                'New Flat Style
                BStyleDown = BDR_SUNKENOUTER
                BStyleUp = BDR_RAISEDINNER
            Case Win95
                'Default Old style
                BStyleUp = BDR_RAISED95
                BStyleDown = BDR_SUNKEN95
            Case Frame
                'Frame style
                BStyleDown = 2
                BStyleUp = 6
        End Select

        If (bState) Then
            'Button Down state
            DrawEdge .hDC, m_CRect, BStyleDown, BF_RECT
        Else
            'Button up State
            DrawEdge .hDC, m_CRect, BStyleUp, BF_RECT
        End If

        'Caption Center
        m_CRect.Top = (yImgPos + 2)
        
        'Text Alignments.
        Select Case m_CapAlign
            Case aLeft
                'Left align.
                If (HasPic) Then
                    m_CRect.Left = (ImgWidth + 6) '+5
                Else
                    m_CRect.Left = 6 '5
                End If
            Case aRight
                'Right align.
               m_CRect.Left = (.ScaleWidth - .TextWidth(m_Caption) - 6) '5
            Case aCenter
                'Center align.
                If (HasPic) Then
                    m_CRect.Left = (.ScaleWidth - .TextWidth(m_Caption) + 6) \ 2 '+5
                Else
                    m_CRect.Left = (.ScaleWidth - .TextWidth(m_Caption)) \ 2
                End If
        End Select
        
        If (HasPic) Then
            'Draw on the picture.
            TransparentBlt .hDC, xImgPos, yImgPos, ImgWidth, ImgWidth, PicImg.hDC, 0, 0, ImgWidth, ImgWidth, RGB(255, 0, 255)
        End If
        
        'Draw the Caption
        DrawTextA .hDC, m_CRect.Left, m_CRect.Top, m_Caption, .Enabled, .ForeColor
    End With
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_CapAlign = PropBag.ReadProperty("CaptionAlignment", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_Caption = PropBag.ReadProperty("Caption", "FlatButton")
    m_ShowRect = PropBag.ReadProperty("ShowFocusRect", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    ButtonStyle = PropBag.ReadProperty("ButtonStyle", 1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CaptionAlignment", m_CapAlign, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, "FlatButton")
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowRect, True)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ButtonStyle", ButtonStyle, 1)
End Sub

Private Sub UserControl_GotFocus()
    m_GotFocus = True
End Sub

Private Sub UserControl_InitProperties()
    m_CapAlign = aCenter
    ButtonStyle = Flat
    m_Caption = "FlatButton"
    m_ShowRect = True
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_LostFocus()
    m_GotFocus = False
    Call UserControl_Paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_Button = Button
    If (m_Button = vbLeftButton) Then
        Call DrawButton(True)
        'Focus
        m_GotFocus = True
        Call BSetFocus
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (m_Button = vbLeftButton) Then
        Call DrawButton
        Call BSetFocus
    End If
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    Call DrawButton
End Sub

Public Property Get CaptionAlignment() As TAlign
    CaptionAlignment = m_CapAlign
End Property

Public Property Let CaptionAlignment(ByVal NewCapAlign As TAlign)
    m_CapAlign = NewCapAlign
    Call DrawButton
    PropertyChanged "CaptionAlignment"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call DrawButton
    PropertyChanged "BackColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call DrawButton
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    Call DrawButton
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_Click()
    If (m_Button <> vbLeftButton) Then
        Exit Sub
    End If
        
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Sub Size(ByVal Width As Single, ByVal Height As Single)
Attribute Size.VB_Description = "Changes the width and height of a User Control."
    UserControl.Size Width, Height
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal vNewCap As String)
    m_Caption = vNewCap
    Call DrawButton
    PropertyChanged "Caption"
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowRect
End Property

Public Property Let ShowFocusRect(ByVal vNewRect As Boolean)
    m_ShowRect = vNewRect
    Call DrawButton
    PropertyChanged "ShowFocusRect"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = PicImg.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set PicImg.Picture = New_Picture
    Call DrawButton
    PropertyChanged "Picture"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call DrawButton
    PropertyChanged "Enabled"
End Property

Public Property Get ButtonStyle() As TButtonStyleA
    ButtonStyle = b_Style
End Property

Public Property Let ButtonStyle(ByVal NewStyle As TButtonStyleA)
    b_Style = NewStyle
    Call DrawButton
    PropertyChanged "ButtonStyle"
End Property
