VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2730
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Options"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox TrayIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
   Begin VB.Menu mnuOP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuCTRL 
         Caption         =   ""
      End
      Begin VB.Menu spe1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###########################################################
' This Form carries a Picture Box named "TrayIcon" set to
' good height so that its not visible when app is running
'###########################################################


'###########################################################
' This Function Show TrayIcon.Picture in System Tray
' Don't Bother what it says
' To change TrayIcon's ToolTip goto 2nd Last Line
'###########################################################
Public Function ShowProgramInTray()
INTRAY = True   'Means App is now in Tray
    
    NI.cbSize = Len(NI) 'set the length of this structure
    NI.hwnd = TrayIcon.hwnd 'control to receive messages from
    NI.uID = 0 'uniqueID
    NI.uID = NI.uID + 1
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP 'operation flags
    NI.uCallbackMessage = WM_MOUSEMOVE 'recieve messages from mouse activities
    NI.hIcon = TrayIcon.Picture  'the location of the icon to display
  
' Change System Tray Icon's Tool Tip Here bt don't delete chr$(0) [its line carriage here]
    
    NI.szTip = "Desktop Manager" + Chr$(0) 'LoadResString(Language) + Chr$(0)  'the tool tip to display"
    result = Shell_NotifyIconA(NIM_ADD, NI) 'add the icon to the system tray
End Function


'###########################################################
' This Function Delete TrayIcon.Picture from System Tray
' Don't Bother what it says
'###########################################################
Private Sub DeleteIcon(pic As Control)
INTRAY = False  'Means app is unloaded or Max mode
    
    ' On remove, we only have to give enough information for Windows
    ' to locate the icon, then tell the system to delete it.
    NI.uID = 0 'uniqueID
    NI.uID = NI.uID + 1
    NI.cbSize = Len(NI)
    NI.hwnd = pic.hwnd
    NI.uCallbackMessage = WM_MOUSEMOVE
    result = Shell_NotifyIconA(NIM_DELETE, NI)
End Sub

'###########################################################
' This Function controls 3 funtions
' 1] Visibility of App Form
' 2] Visibility of Tray Icon
' 3] Menu Caption Control
'###########################################################

Public Function NoSysIcon(maxIcon As Boolean)
    Select Case maxIcon
    Case False   'Case App in Min Mode
        Me.Visible = False
        ShowProgramInTray               'Now show TrayIcon PictureBox's Picture in SysTray as icon
        mnuCTRL.Caption = "E&xpand Application"
    
    Case Else   'Case App in Max Mode
        Me.Visible = True
        DeleteIcon TrayIcon
        mnuCTRL.Caption = "Minimize App to System Tray"
    
    End Select

End Function


Private Sub cmdMenu_Click()
PopupMenu mnuOP, 2, cmdMenu.Left + cmdMenu.Width / 2, cmdMenu.Top + cmdMenu.Height / 2, mnuCTRL 'show the popoup menu
End Sub

Private Sub Form_Load()
TrayIcon.Top = Me.Height + 1000 'Set TrayIcon PictureBox top to such limit that its not visible

NoSysIcon False
'Initially we are setting App to Min(or minimized in systray) mode
'So this must be initially false
                             
'Cooool Funda used here:
'App visible = True when Tray Icon Visible = False And Vice-versa

End Sub

'mnuCTRL controls the Expansion and Minimizing function
Private Sub mnuCTRL_Click()
NoSysIcon INTRAY  'Call Menu Function
End Sub

Private Sub mnuExit_Click()
DeleteIcon TrayIcon 'As we are exiting App
End
End Sub



'###########################################################
' This Function Tells what Event happened to icon in
' system tray
'###########################################################
Private Sub Trayicon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Msg As Long
    Msg = (X And &HFF) * &H100

    Select Case Msg
        Case 0 'mouse moves
        
        Case &HF00  'left mouse button down
        
        Case &H1E00 'left mouse button up
        
        Case &H3C00  'right mouse button down
        PopupMenu mnuOP, 2, , , mnuCTRL 'show the popoup menu
        
        Case &H2D00 'left mouse button double click
        NoSysIcon True    'Show App on double clicking Mouse's Left Button
        
        Case &H4B00 'right mouse button up
        
        Case &H5A00 'right mouse button double click
        
    End Select
   
End Sub

