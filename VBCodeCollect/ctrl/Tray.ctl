VERSION 5.00
Begin VB.UserControl Tray 
   Appearance      =   0  'Flat
   BackColor       =   &H80000002&
   CanGetFocus     =   0   'False
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   330
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   330
   ScaleWidth      =   330
End
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    UID As Long
    uFlags As Long
    uCallBackmessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_MOUSEOVER = &H200

Private Declare Function Shell_NotifyIcon Lib "shell32" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Dim nid As NOTIFYICONDATA

Const m_def_Visible = False
Const m_def_ToolTip = ""
Dim m_Visible As Boolean
Dim m_ToolTip As String
Dim m_Icon As Picture

Event MouseMove()
Event MouseDown(Button As Integer)
Event MouseUp(Button As Integer)
Event DblClick()

Public Property Get Icon() As Picture
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set m_Icon = New_Icon
    If New_Icon Is Nothing Then
        Visible = False
    Else
        If m_Visible Then
            nid.uFlags = NIF_ICON
            nid.hIcon = m_Icon
            Shell_NotifyIcon NIM_MODIFY, nid
        End If
    End If
    PropertyChanged "Icon"
End Property

Public Property Get ToolTip() As String
    ToolTip = m_ToolTip
End Property

Public Property Let ToolTip(ByVal New_ToolTip As String)
    m_ToolTip = Trim(New_ToolTip)
    nid.uFlags = NIF_TIP
    nid.szTip = m_ToolTip & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, nid
    PropertyChanged "ToolTip"
End Property

Private Sub UserControl_InitProperties()
    Set m_Icon = LoadPicture("")
    m_ToolTip = m_def_ToolTip
    m_Visible = m_def_Visible
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    m_ToolTip = PropBag.ReadProperty("ToolTip", m_def_ToolTip)
    m_Visible = PropBag.ReadProperty("Visible", m_def_Visible)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Size 300, 300
    If Err Then Err.Clear
End Sub

Private Sub UserControl_Terminate()
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("ToolTip", m_ToolTip, m_def_ToolTip)
    Call PropBag.WriteProperty("Visible", m_Visible, m_def_Visible)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case X / Screen.TwipsPerPixelX
    Case WM_LBUTTONDBLCLK
        RaiseEvent DblClick
    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(vbLeftButton)
    Case WM_LBUTTONUP
        RaiseEvent MouseUp(vbLeftButton)
    Case WM_RBUTTONDBLCLK
        RaiseEvent DblClick
    Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(vbRightButton)
    Case WM_RBUTTONUP
        RaiseEvent MouseUp(vbRightButton)
    Case WM_MOUSEOVER
        RaiseEvent MouseMove
    End Select
End Sub

Public Property Get Visible() As Boolean
Attribute Visible.VB_MemberFlags = "400"
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal New_Visible As Boolean)
    If m_Visible = New_Visible Then Exit Property
    m_Visible = New_Visible
    If m_Visible Then
        If Ambient.UserMode Then
            nid.cbSize = Len(nid)
            nid.hwnd = UserControl.hwnd
            nid.UID = Int((Rnd * 65535) + 1)
            nid.uFlags = NIF_MESSAGE
            If Not m_Icon Is Nothing Then
                nid.uFlags = nid.uFlags + NIF_ICON
                nid.hIcon = m_Icon
            End If
            If m_ToolTip <> "" Then
                nid.uFlags = nid.uFlags + NIF_TIP
                nid.szTip = m_ToolTip & vbNullChar
            End If
            nid.uCallBackmessage = WM_MOUSEMOVE
            Shell_NotifyIcon NIM_ADD, nid
        End If
    Else
        Shell_NotifyIcon NIM_DELETE, nid
    End If
    PropertyChanged "Visible"
End Property

