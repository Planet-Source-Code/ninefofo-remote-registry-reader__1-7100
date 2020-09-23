Attribute VB_Name = "GetConnected"
Option Explicit

Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long

Type NETRESOURCE
dwScope As Long
dwType As Long
dwDisplayType As Long
dwUsage As Long
lpLocalName As String
lpRemoteName As String
lpComment As String
lpProvider As String
End Type

Type LANConnections
Path As String
UserName As String
Password As String
End Type

Public Function GetDrive(Drive As String, Path As String, UserName As String, Password As String) As Long

Dim MyNetStruct As NETRESOURCE
    MyNetStruct.dwType = 0
    MyNetStruct.lpLocalName = Drive & Chr$(0)
    MyNetStruct.lpRemoteName = Path & Chr$(0)
    MyNetStruct.lpProvider = ""
GetDrive = WNetAddConnection2(MyNetStruct, Password & Chr$(0), UserName & Chr$(0), 0)

End Function
Public Function DisDrive(Drive As String) As Long
    DisDrive = WNetCancelConnection2(Drive & Chr$(0), 0, 1)
End Function
Public Sub GetIPCConnection(systemip As String)

Dim Maps As LANConnections
Dim AResult As Long
Dim ZResult As String
Dim MsgStr As String

Maps.Path = "\\" + Trim(systemip) + "\ipc$"
Maps.UserName = Trim(frmLogin.txtUserName.Text)
Maps.Password = Trim(frmLogin.txtPassword.Text)

AResult = GetDrive("", Maps.Path, Maps.UserName, Maps.Password)

If AResult = 0 Then Exit Sub Else
GoTo ErrorDetected

ErrorDetected:
ZResult = FriendlyError(AResult)
    MsgBox ZResult, vbCritical, "GetIPCConnection"

Exit Sub

End Sub
Public Sub DisIPCConnection(systemip As String)
Dim Maps As LANConnections
Dim AResult As Long
Dim ZResult As String

Maps.Path = "\\" + Trim(systemip) + "\ipc$"

AResult = DisDrive(Maps.Path)

If AResult = 0 Then Exit Sub Else
GoTo ErrorDetected

ErrorDetected:
ZResult = FriendlyError(AResult)
    MsgBox ZResult, vbCritical, "DisIPCConnection"

Exit Sub

End Sub
