VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynDnsSmartUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type DynDnsDef
    Username    As String
    Password    As String
    InetIP     As String
    Hostname    As String
    MX          As String
    Wildcard    As Boolean
    BackMX      As Boolean
    Offline     As Boolean
End Type

Private DynDnsDef As DynDnsDef

Public Property Let Username(ByVal sUsername As String)
    DynDnsDef.Username = sUsername
End Property

Public Property Get Username() As String
    Username = DynDnsDef.Username
End Property

Public Property Let Password(ByVal sPassword As String)
    If Len(sPassword) > 0 Then
        DynDnsDef.Password = Crypt(sPassword)
    Else
        DynDnsDef.Password = vbNullString
    End If
End Property

Public Property Get Password() As String
    If Len(DynDnsDef.Password) Then
        Password = Crypt(DynDnsDef.Password)
    Else
        Password = vbNullString
    End If
End Property

Public Property Let InetIP(ByVal sInetIP As String)
    DynDnsDef.InetIP = sInetIP
End Property

Public Property Get InetIP() As String
    InetIP = DynDnsDef.InetIP
End Property

Public Property Let Hostname(ByVal sHostname As String)
    DynDnsDef.Hostname = sHostname
End Property

Public Property Get Hostname() As String
    Hostname = DynDnsDef.Hostname
End Property

Public Property Get MXServer() As String
    MXServer = DynDnsDef.MX
End Property

Public Property Let MXServer(ByVal sMXServer As String)
    DynDnsDef.MX = sMXServer
End Property

Public Property Let Wildcard(ByVal bWildcard As Boolean)
    DynDnsDef.Wildcard = bWildcard
End Property

Public Property Get Wildcard() As Boolean
    Wildcard = DynDnsDef.Wildcard
End Property

Public Property Let MXServerBackup(ByVal bBackupMX As Boolean)
    DynDnsDef.BackMX = bBackupMX
End Property

Public Property Get MXServerBackup() As Boolean
    MXServerBackup = DynDnsDef.BackMX
End Property

Public Property Let Offline(ByVal bOffline As Boolean)
    DynDnsDef.Offline = bOffline
End Property

Public Property Get Offline() As Boolean
    Offline = DynDnsDef.Offline
End Property

Public Function DynDnsUpdate() As String
On Error GoTo ErrH
    
    Dim hXML As Object
    Set hXML = CreateObject("microsoft.xmlhttp")
    
    Call hXML.Open("get", UpdateString, False)
    Call hXML.setrequestheader("Pragma", "no-cache")
    Call hXML.setrequestheader("Cache-control", "no-cache")
    
    Call hXML.send
    
    DynDnsUpdate = hXML.responsetext
    
ErrH: If Err Then DynDnsUpdate = "Update failed! " & Chr(13) & Err.Description
End Function

Private Function GetBool(ByVal Bool As Boolean, Optional ByVal Typ As Integer = 0) As String

    Select Case Typ
        Case 0
            If Bool Then
                GetBool = "ON"
            Else
                GetBool = "OFF"
            End If
        Case 1
            If Bool Then
                GetBool = "YES"
            Else
                GetBool = "NO"
            End If
    End Select
    
End Function

Public Property Get UpdateString() As String
On Error GoTo ErrH
  
    Dim updStr As String
    updStr = "http://" & DynDnsDef.Username
    updStr = updStr & ":" & Me.Password
    updStr = updStr & "@members.dyndns.org/nic/update?system=dyndns&hostname="
    updStr = updStr & DynDnsDef.Hostname & "&myip=" & DynDnsDef.InetIP
    updStr = updStr & "&wildcard=" & GetBool(DynDnsDef.Wildcard) & "&mx=" & DynDnsDef.MX
    updStr = updStr & "&backmx=" & GetBool(DynDnsDef.BackMX, 1) & "&offline=" & GetBool(DynDnsDef.Offline, 1)
    
    UpdateString = updStr
    
ErrH:
End Property

Public Function LoadSettings(ByVal filename As String) As Boolean
On Error GoTo ErrH

    If Dir(filename, vbNormal) <> vbNullString Then
        
        Dim fNR As Integer: fNR = FreeFile
        Open filename For Binary As #fNR
        Get #fNR, , DynDnsDef
        Close #fNR
        
        LoadSettings = True
        
    Else
        
        LoadSettings = False
    
    End If

ErrH: If Err Then LoadSettings = False
End Function

Public Function SaveSettings(ByVal filename As String) As Boolean
On Error GoTo ErrH

    Dim fNR As Integer: fNR = FreeFile
    Open filename For Binary As #fNR
    Put #fNR, , DynDnsDef
    Close #fNR
    SaveSettings = True

ErrH: If Err Then SaveSettings = False
End Function

Private Function Crypt(ByVal Expression As String) As String
On Error GoTo ErrH:

    Dim CrStr   As String, _
        i       As Long, _
        CrAsc   As Byte
    
    If Len(Expression) > 0 Then
        
        For i = Len(Expression) To 1 Step -1
            CrAsc = Asc(Mid(Expression, i, 1))
            If CrAsc = 255 Then CrAsc = 254
            CrStr = CrStr & Chr(Not CrAsc)
        Next
        
        Crypt = CrStr
    
    End If

ErrH: If Err Then Crypt = Expression
End Function

Public Property Get CurrentInetIP() As String
On Error GoTo ErrH

    Dim hXML    As Object, _
        retStr  As String, _
        semPos  As Integer, _
        addPos  As Integer
    
    Set hXML = CreateObject("microsoft.xmlhttp")

    Call hXML.Open("get", "http://checkip.dyndns.org/", False)
    Call hXML.setrequestheader("Pragma", "no-cache")
    Call hXML.setrequestheader("Cache-control", "no-cache")

    Call hXML.send
    
    retStr = hXML.responsetext
    semPos = InStr(1, retStr, "Current IP Address: ", vbTextCompare)
    addPos = semPos + 20
    
    If semPos > 0 Then
        retStr = Mid(retStr, addPos)
        semPos = InStr(1, retStr, "<br>", vbTextCompare) - 1
        If semPos > 0 Then retStr = Mid(retStr, 1, semPos)
        If InStr(1, retStr, Chr(10)) > 0 Then retStr = Mid(retStr, 1, Len(retStr) - 1)
        
    End If
    
    CurrentInetIP = Trim(retStr)
    
ErrH: If Err Then CurrentInetIP = "IP Detection failed! " & Chr(13) & Err.Description
End Property

Public Sub About()
    Call MsgBox("DynDns Smartupdate is written by Frank Diehl (c)2003" & Chr(13) & "eMail to: frank-diehl@gmx.de", vbInformation, "About")
End Sub

