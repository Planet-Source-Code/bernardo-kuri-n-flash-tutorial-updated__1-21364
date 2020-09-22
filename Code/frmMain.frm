VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3720
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash xWin 
      Height          =   3750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3750
      _cx             =   4200919
      _cy             =   4200919
      Movie           =   ""
      Src             =   ""
      WMode           =   "Transparent"
      Play            =   -1  'True
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ReleaseCapture& Lib "user32" ()
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)
Private Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Const HTCAPTION As Long = 2
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Sub CreateMovie()
Dim ByteArray() As Byte, sEnv$

sEnv = Environ("TEMP")

If MovieExists Then Kill sEnv & "\tutorial.swf"

ByteArray = LoadResData(555, "SWF")
Open sEnv & "\tutorial.swf" For Binary Access Write As #1
    Put #1, , ByteArray()
Close #1

xWin.Movie = sEnv & "\tutorial.swf"
End Sub
Private Function MovieExists() As Boolean
MovieExists = (Dir(Environ("TEMP") & "\tutorial.swf") <> "")
End Function
Private Sub Form_Load()
CreateMovie
OnTop Me, True
End Sub
Private Sub Form_Resize()
Select Case Me.WindowState
    Case vbMinimized
        Caption = "Flash in VB Tutorial - by Bernardo Kuri N."
    
    Case vbNormal
        Caption = ""
        Move Left, Top, 3750, 3750
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
OnTop Me, False
Set frmMain = Nothing
End Sub
Private Sub xWin_FSCommand(ByVal command As String, ByVal args As String)
Static NhWnd& 'Variable to store WordPad's window handle

Select Case command
    Case "tb_action"
        Select Case args
            Case "move_win"
                ReleaseCapture
                SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

            Case "close_win"
                If MsgBox("Are you sure you want to close this tutorial?", vbYesNo + vbInformation + vbDefaultButton2, "End Tutorial?") = vbYes Then End
            
            Case "min_win"
                Me.WindowState = vbMinimized
                
        End Select
    Case "win_action"
        Select Case args
            Case "open_notepad" 'Opens NotePad
                Shell "C:\WINNT\notepad.exe", vbNormalFocus
                NhWnd = FindWindow(vbNullString, "Untitled - NotePad")
                Call SetWindowText(NhWnd, "Flash Tutorial - NotePad")
            
            Case "write_text" 'Writes Text into NotePad
                Clipboard.Clear
                Clipboard.SetText InputBox("Which text do you want to write into NotePad?", "Write Text Into NotePad", "Visual Basic integration into Flash -- A Simple Tutorial by Bernardo Kuri N."), 1
                AppActivate "Flash Tutorial - NotePad"
                SendKeys "%EP"
                
            Case "close_notepad" 'Closes NotePad
                Clipboard.Clear
                AppActivate "Flash Tutorial - NotePad"
                SendKeys "%{F4}N"
        End Select
End Select
End Sub
Private Sub OnTop(ByVal frm As Form, ByVal YesNo As Boolean)
Dim TopPixels%, LeftPixels%, WidthPixels%, HeightPixels%

TopPixels = frm.Top / Screen.TwipsPerPixelY
LeftPixels = frm.Left / Screen.TwipsPerPixelX
WidthPixels = frm.Width / Screen.TwipsPerPixelY
HeightPixels = frm.Height / Screen.TwipsPerPixelX

Select Case YesNo
    Case True: SetWindowPos frm.hwnd, -1, LeftPixels, TopPixels, WidthPixels, HeightPixels, &H50
    Case False: SetWindowPos frm.hwnd, -2, LeftPixels, TopPixels, WidthPixels, HeightPixels, &H50
End Select
End Sub
