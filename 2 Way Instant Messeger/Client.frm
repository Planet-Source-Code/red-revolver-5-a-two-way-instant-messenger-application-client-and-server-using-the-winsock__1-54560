VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Client 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client 2"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "Client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wskclient 
      Left            =   4200
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtp 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtrecieve 
      Height          =   1965
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtsend 
      Height          =   645
      Left            =   960
      TabIndex        =   6
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox txtip 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   90
      Width           =   3615
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "Send"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Recieving:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Sending:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Server IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub cmdconnect_Click()
On Error Resume Next
Dim ip As String
Dim p As String
ip = txtip.Text
p = txtp.Text
wskclient.Connect ip, p
MsgBox "Connected to Server", vbInformation, "Connected"
End Sub

Private Sub cmdsend_Click()
On Error Resume Next
Dim data As String
data = txtsend.Text
wskclient.SendData data
txtsend.Text = ""
End Sub

Private Sub Form_Load()
txtp.Text = "5432"
txtip.Text = "127.0.0.1"
End Sub

Private Sub wskclient_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim data As String
wskclient.GetData data
txtrecieve.Text = data
End Sub

Private Sub wskclient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim data As String
wskclient.GetData data
txtrecieve.Text = data
End Sub
