Attribute VB_Name = "modTrayIcon"
Option Explicit
' Thanks to David Carta
' This code is based on nice work done by Wolf.
' By: David Carta
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=44749&lngWId=1

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
    
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    'Make your own constant, e.g.:
    Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_LBUTTONDBLCLK = &H203
    Public Const WM_LBUTTONDOWN = &H201
    Public Const WM_RBUTTONDOWN = &H204
    Private m_strTip$
    Private TaskBarIconJustClicked As Boolean
    
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias _
    "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As _
    NOTIFYICONDATA) As Long
'ststray subs  By: David Carta
    
Public Sub CreateIcon(PictureBox As Control)
    Dim Tic As NOTIFYICONDATA
    Dim strTip As String
    Dim erg As Boolean
    SetToolTip "Auto File Copy"
    Tic.cbSize = Len(Tic)
    Tic.hwnd = PictureBox.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = PictureBox.Picture
    Tic.szTip = strTip & Chr$(0)
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Public Sub DeleteIcon(PictureBox As Control)
    Dim Tic As NOTIFYICONDATA
    Dim erg As Boolean
    Tic.cbSize = Len(Tic)
    Tic.hwnd = PictureBox.hwnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Public Sub SetToolTip(strTip As String)
    m_strTip$ = strTip
End Sub

Public Sub MoveToTaksbar(frm As Form, Pic1 As PictureBox)
    Static RunningThisFunction As Boolean
    
    If RunningThisFunction = True Then Exit Sub
    RunningThisFunction = True


    If (TaskBarIconJustClicked = True) Then
        frm.WindowState = vbNormal
        frm.Show
        TaskBarIconJustClicked = False
        RunningThisFunction = False
        Exit Sub
    End If
    


    If (frm.WindowState = vbMinimized) Then
        frm.Hide
        Call CreateIcon(Pic1)
    End If
    RunningThisFunction = False
End Sub

Public Sub RestoreFromTaskbar(frm As Form, Pic1 As PictureBox, Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
    X = X / Screen.TwipsPerPixelX


    Select Case X
        Case WM_LBUTTONDOWN
        Case WM_RBUTTONDOWN
        Case WM_MOUSEMOVE
        Case WM_LBUTTONDBLCLK
        Call DeleteIcon(Pic1)
        TaskBarIconJustClicked = True
        frm.Show
    End Select
End Sub


