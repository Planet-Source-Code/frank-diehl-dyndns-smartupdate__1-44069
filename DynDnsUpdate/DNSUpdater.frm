VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Smartupdate (non-commercial)"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "DNSUpdater.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3510
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picDynDns 
      Align           =   1  'Oben ausrichten
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      MouseIcon       =   "DNSUpdater.frx":6852
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "DNSUpdater.frx":6B5C
      ScaleHeight     =   900
      ScaleWidth      =   3510
      TabIndex        =   17
      Top             =   0
      Width           =   3510
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   1005
      TabIndex        =   8
      Top             =   3285
      Width           =   870
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   135
      TabIndex        =   7
      Top             =   3285
      Width           =   870
   End
   Begin VB.TextBox txtPwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1305
      Width           =   2055
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   1005
      Width           =   2055
   End
   Begin VB.TextBox txtMXServer 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2625
      Width           =   2055
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Now"
      Height          =   375
      Left            =   2115
      TabIndex        =   9
      Top             =   3285
      Width           =   1275
   End
   Begin VB.CheckBox chkBackUpXM 
      Caption         =   "Check1"
      Height          =   240
      Left            =   1320
      TabIndex        =   6
      Top             =   2940
      Width           =   210
   End
   Begin VB.CheckBox chkWildcard 
      Caption         =   "Check1"
      Height          =   240
      Left            =   1320
      TabIndex        =   4
      Top             =   2355
      Width           =   210
   End
   Begin VB.TextBox txtHostname 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1995
      Width           =   2055
   End
   Begin VB.TextBox txtInetIP 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1695
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   225
      Left            =   165
      TabIndex        =   16
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   225
      Left            =   165
      TabIndex        =   15
      Top             =   1020
      Width           =   1260
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Backup MX"
      Height          =   225
      Left            =   165
      TabIndex        =   14
      Top             =   2955
      Width           =   1260
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Mail Exchanger"
      Height          =   225
      Left            =   165
      TabIndex        =   13
      Top             =   2655
      Width           =   1260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Wildcard"
      Height          =   225
      Left            =   165
      TabIndex        =   12
      Top             =   2355
      Width           =   1260
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Hostname"
      Height          =   225
      Left            =   165
      TabIndex        =   11
      Top             =   2025
      Width           =   1260
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "IP - Address"
      Height          =   225
      Left            =   165
      TabIndex        =   10
      Top             =   1740
      Width           =   1260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private DynDns As New DynDnsSmartUpdate

Private Sub cmdAbout_Click()
    Call DynDns.About
End Sub

Private Sub cmdExit_Click()
    Call UpdateSettings
    End
End Sub

Private Sub cmdUpdate_Click()
    
    Call UpdateSettings
    Call MsgBox("DynDns ReturnCode: " & DynDns.DynDnsUpdate, vbInformation, "Update State")
       
End Sub

Private Sub Form_Load()

    With DynDns
    
        Call .LoadSettings(App.Path & "\DynDns.inf")
        
        Me.txtUsername = .Username
        Me.txtPwd = .Password
        Me.txtHostname = .Hostname
        Me.txtInetIP = .CurrentInetIP
        Me.txtMXServer = .MXServer
        Me.chkBackUpXM = -.MXServerBackup
        Me.chkWildcard = -.Wildcard
                
    End With


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call UpdateSettings
    
End Sub

Private Sub UpdateSettings()

    With DynDns
        
        .Username = Me.txtUsername
        .Password = Me.txtPwd
        .Hostname = Me.txtHostname
        .InetIP = Me.txtInetIP
        .MXServer = Me.txtMXServer
        .MXServerBackup = -Me.chkBackUpXM
        .Wildcard = -Me.chkWildcard
                
        Call .SaveSettings(App.Path & "\DynDns.inf")
        
    End With

End Sub

Private Sub picDynDns_Click()
    
    Call ShellExecute(Me.hWnd, "Open", "http:\\www.dyndns.org", vbNullString, vbNullString, 1)
    
End Sub
