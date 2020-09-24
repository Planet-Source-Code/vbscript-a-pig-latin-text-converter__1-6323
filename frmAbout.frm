VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "About AboutForm"
   Begin VB.TextBox txtCopyright 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmAbout.frx":0E42
      Top             =   2040
      Width           =   3855
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   690
      Left            =   120
      Picture         =   "frmAbout.frx":0E98
      ScaleHeight     =   630
      ScaleMode       =   0  'User
      ScaleWidth      =   630
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   690
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2040
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Tag             =   "&System Info..."
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   1050
      Left            =   1050
      TabIndex        =   4
      Tag             =   "App Description"
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label lblTitle 
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   3
      Tag             =   "Application Title"
      Top             =   240
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225
      X2              =   5657
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const KEY_ALL_ACCESS = &H2003F
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    frmAbout.Caption = App.Title & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = App.Comments
End Sub

Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
        Unload Me
        frmMain.Show
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
        Dim rc As Long
        Dim SysInfoPath As String
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                Else
                        GoTo SysInfoErr
                End If
        Else
                GoTo SysInfoErr
        End If
        Call Shell(SysInfoPath, vbNormalFocus)
        Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long
        Dim rc As Long
        Dim hKey As Long
        Dim hDepth As Long
        Dim KeyValType As Long
        Dim tmpVal As String
        Dim KeyValSize As Long
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
        tmpVal = String$(1024, 0)
        KeyValSize = 1024
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        Select Case KeyValType
        Case REG_SZ
                KeyVal = tmpVal
        Case REG_DWORD
                For i = Len(tmpVal) To 1 Step -1
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
                Next
                KeyVal = Format$("&h" + KeyVal)
        End Select
        GetKeyValue = True
        rc = RegCloseKey(hKey)
        Exit Function
GetKeyError:
        KeyVal = ""
        GetKeyValue = False
        rc = RegCloseKey(hKey)
End Function
