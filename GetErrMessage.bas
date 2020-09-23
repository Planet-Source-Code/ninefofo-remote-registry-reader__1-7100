Attribute VB_Name = "GetErrMessage"
Option Explicit

Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Function FriendlyError(ErrNo As Long) As String

Dim res As Long, Buffer As String * 256
res = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, ErrNo, 0, Buffer, Len(Buffer), 0)
If res > 0 Then
    FriendlyError = Left(Buffer, res - 2) & " (" & Format(ErrNo) & ")"
Else
    FriendlyError = "Error number " & Format(ErrNo)
End If
End Function

