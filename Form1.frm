VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Online Status"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   0
      Left            =   4200
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   1
      Left            =   4800
      Picture         =   "Form1.frx":0614
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer tmrSysTray 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4800
      Top             =   120
   End
   Begin VB.CheckBox chkMinToTray 
      Caption         =   "Minimize to System Tray"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton cmdOnline 
      Caption         =   "Online"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOffline 
      Caption         =   "Offline"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4320
      Top             =   120
   End
   Begin VB.Menu mPopUpMenu 
      Caption         =   "Work Offline Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuOffOnLine 
         Caption         =   "Work Offline"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSepOne 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////
'Created by George Selos
'Thanks to Peter Verburgh for idea on destroying pop up menu
'Thanks also to E.Spencer for Registry Access module
'Any comments email me at gchelos@hotmail.com
'This is Freeware.
'///////////////////////////////////////

'Use this to quickly work Online or Offline
'when you click on the icon in System Tray
'Icon also changes when another program changes the Online Status
'Has been tested on Win 2000 only
'///////////////////////////////////////

Dim nidTemp As NOTIFYICONDATA

Dim CurrentToolTip As String
Dim MnuCaption As String

Dim j As Long, trayIconNum As Long
Dim hMenu As Long
Dim Msg As Long

Dim currstate As Long
Dim AreWeOffline As Long

Dim mSh As Long
Dim HndSysWnd As Long
Dim HndMouPos As Long

Dim Pt As POINTAPI
Dim Insystray As Boolean

Private Sub chkMinToTray_Click()
If chkMinToTray.Value = 0 Then
 Delete_TrayIcon
 cmdHide.Enabled = False
Else
 AddModify_TrayIcon currstate
 cmdHide.Enabled = True
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHide_Click()
Me.Hide
End Sub

Private Sub cmdOffline_Click()
GoOffline True
End Sub

Private Sub cmdOnline_Click()
GoOffline False
End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then Unload Me: Exit Sub
Insystray = False
trayIconNum = -1
currstate = CLng(ReadRegistry(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "GlobalUserOffline"))
AddModify_TrayIcon currstate
Me.Hide
tmrSysTray.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Destroy_Menu
 Delete_TrayIcon
 End
End Sub

Private Sub Timer1_Timer()

'After 600ms if the cursor is outside the system menu -
'pop up menu form will be destroyed
'Note: Must also unload DummyForm

'Debug.Print "Timer1"
On Error Resume Next
GetCursorPos Pt
'get handle based on mouse position
HndMouPos = WindowFromPointXY(Pt.X, Pt.Y)
HndSysWnd = FindWindow("#32768", vbNullString)
If HndMouPos = GetTrayhWnd Then 'cursor over system tray icons
 mSh = 0
ElseIf HndSysWnd <> HndMouPos Then 'if cursor outside pop up menu
 'wait 600ms
 mSh = mSh + 1
 If mSh > 1 Then
  Destroy_Menu
  Timer1.Enabled = False
 End If
Else 'cursor over pop up menu
 mSh = 0
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.ScaleMode = vbPixels Then
     Msg = X
    Else
     Msg = X / Screen.TwipsPerPixelX
    End If

    Select Case Msg
    Case WM_LBUTTONUP '514
     If hMenu <> 0 Then Destroy_Menu: Exit Sub
     ChangeOnlineStatus
    Case WM_LBUTTONDBLCLK '515
    
    Case WM_RBUTTONUP '517 display popup menu
     Create_PopUp_Menu
    End Select
    
End Sub

Private Sub tmrSysTray_Timer()

'check if another program changes online status
currstate = CLng(ReadRegistry(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "GlobalUserOffline"))

If chkMinToTray.Value = 1 Then
 If Me.WindowState = vbMinimized Then Me.Hide
 If currstate <> AreWeOffline Then AddModify_TrayIcon currstate
Else
 If currstate <> AreWeOffline Then Form_Icon
End If

AreWeOffline = currstate
End Sub

Public Sub Form_Icon()

If currstate = 0 Then
 Me.Icon = LoadPicture(App.Path & "\Online.ico")
 Me.Caption = "We are working Online"
Else
 Me.Icon = LoadPicture(App.Path & "\Offline.ico")
 Me.Caption = "We are working Offline"
End If

'Debug.Print Me.Icon
End Sub

Public Sub AddModify_TrayIcon(picIndex As Long)

'BETTER TO USE USE PICTUREBOX TO ADD AN ICON OR IMAGE TO SYSTEM TRAY
If trayIconNum = picIndex And Insystray = True Then Exit Sub

Dim mTip$
Dim mIcon As PictureBox
Set mIcon = picIcon(picIndex)

'modify form's icon property
If picIndex = 0 Then
 Me.Icon = LoadPicture(App.Path & "\Online.ico")
 Me.Caption = "We are working Online"
 MnuCaption = "Work Offline"
 mTip = "Online"
Else
 Me.Icon = LoadPicture(App.Path & "\Offline.ico")
 Me.Caption = "We are working Offline"
 MnuCaption = "Work Online"
 mTip = "Offline"
End If
  
nidTemp.cbSize = Len(nidTemp)
nidTemp.hWnd = Me.hWnd
nidTemp.uId = vbNull
nidTemp.uFlags = NIF_DOALL
nidTemp.uCallBackMessage = WM_MOUSEMOVE
nidTemp.hIcon = mIcon.Picture
nidTemp.szTip = mTip & Chr$(0)

trayIconNum = picIndex
CurrentToolTip = mTip
    
  If Insystray = False Then
   j = Shell_NotifyIcon(NIM_ADD, nidTemp)
   Insystray = True
  Else
   j = Shell_NotifyIcon(NIM_MODIFY, nidTemp)
  End If

'Debug.Print "AddModify_TrayIcon"
End Sub

Public Sub Delete_TrayIcon()
If Insystray = True Then
 j = Shell_NotifyIcon(NIM_DELETE, nidTemp)
 trayIconNum = -1
 Insystray = False
End If
End Sub

Public Sub ChangeOnlineStatus()

If MnuCaption = "Work Offline" Then
 GoOffline True
 AddModify_TrayIcon 1
Else
 GoOffline False
 AddModify_TrayIcon 0
End If

End Sub

Public Sub Create_PopUp_Menu()

Dim xId As Long
Dim mRetMenu As Long

If hMenu <> 0 Then Exit Sub
mSh = 0
Timer1.Enabled = True
hMenu = CreatePopupMenu()

'Append a few menu items
AppendMenu hMenu, MF_STRING, "1000", MnuCaption
AppendMenu hMenu, MF_STRING, "2000", "Restore"
AppendMenu hMenu, MF_SEPARATOR, "3000", ByVal 0&
AppendMenu hMenu, MF_STRING, "4000", "Exit"

GetCursorPos Pt 'Get the position of the mouse cursor
Load frmDummy 'this invisible window receives all messages from the pop up menu
mRetMenu = TrackPopupMenu(hMenu, TPM_RETURNCMD, Pt.X, Pt.Y, 0, frmDummy.hWnd, ByVal 0&)

Select Case mRetMenu
Case 0
 Destroy_Menu
Case 1000
 ChangeOnlineStatus
 Destroy_Menu
Case 2000
 ShowWindow Me.hWnd, SW_RESTORE
 Me.Left = (Screen.Width / 2) - (Me.Width / 2)
 Me.Top = (Screen.Height / 2) - (Me.Height / 2)
 SetForegroundWindow Me.hWnd
 Destroy_Menu
Case 4000
 Unload Me
End Select

End Sub

Public Sub Destroy_Menu()
If hMenu <> 0 Then
 DestroyMenu hMenu
 hMenu = 0
 Unload frmDummy
End If
End Sub

Public Function GetTrayhWnd() As Long

Dim OurParent As Long
Dim OurFirstChild As Long
Dim OurSysTray As Long

'handle to the shell tray
OurParent = FindWindow("Shell_TrayWnd", "")
'handle to the clock
OurFirstChild = FindWindowEx(OurParent, 0&, "TrayNotifyWnd", vbNullString)
'handle to the area where the icons reside
OurSysTray = FindWindowEx(OurFirstChild, 0&, "ToolbarWindow32", vbNullString)

GetTrayhWnd = OurSysTray

End Function

