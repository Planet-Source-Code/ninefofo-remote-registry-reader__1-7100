VERSION 5.00
Begin VB.Form frmRemoteRead 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Remote Peer Registry"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   5415
   End
   Begin VB.Frame Frame5 
      Caption         =   "Computer Name/IP Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtSystemName 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Read Remote Registry Value"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   5415
   End
   Begin VB.Frame Frame4 
      Caption         =   "Returned Value"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   5415
      Begin VB.TextBox txtRetVal 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Read Value"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   5415
      Begin VB.TextBox txtValueString 
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Text            =   "ProductId"
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Open Hive"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5415
      Begin VB.TextBox txtHiveString 
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Text            =   "HKEY_LOCAL_MACHINE"
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Open Key"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   5415
      Begin VB.TextBox txtKeyString 
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Text            =   "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
         Top             =   330
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmRemoteRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExec_Click()
frmLogin.Show 1
DoEvents
PeerRegistry
Beep
End Sub

Private Sub PeerRegistry()
Dim OpenKeyVal As Long
Dim RResult As Long
Dim OpenHiveVal As Long
Dim StrValue, CMPTRName, keytogo As String, InfoTextStr As String
Dim x As Integer

GetIPCConnection (Trim(txtSystemName.Text))

CMPTRName = Trim(txtSystemName.Text)

RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
keytogo = Trim(txtKeyString.Text)
OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
InfoTextStr = RegistryQueryValue(OpenKeyVal, Trim(txtValueString), REG_SZ)

RegCloseKey (OpenKeyVal)
RegCloseKey (OpenHiveVal)

DisIPCConnection (Trim(txtSystemName.Text))

txtRetVal.Text = Trim(InfoTextStr)

txtSystemName.SetFocus
txtSystemName.SelStart = 0
txtSystemName.SelLength = Len(txtSystemName.Text)

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
txtSystemName.Text = Environ("USERDOMAIN")
txtSystemName.SelStart = 0
txtSystemName.SelLength = Len(txtSystemName.Text)
End Sub

