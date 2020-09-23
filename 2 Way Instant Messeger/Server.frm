VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server (Client 1)"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "Send"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtsend 
      Height          =   645
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox txtrecieve 
      Height          =   1935
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin MSWinsockLib.Winsock wskserver 
      Left            =   120
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Sending:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Recieving:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsend_Click()
On Error Resume Next
Dim data As String
data = txtsend.Text
wskserver.SendData data
txtsend.Text = ""
End Sub

Private Sub Command1_Click()
MsgBox "This program was designed and programmed by Red Revolver 5. Please leave comments and vote for me at PSC.", vbInformation, "About"
End Sub

Private Sub Form_Load()
Dim p As String
p = 5432
wskserver.LocalPort = p
wskserver.Listen
MsgBox "Server Listening. Now open the client and press connect. If you experience errors, please make sure that the IP address is set to 127.0.0.1 and the port is set to 5432.", vbInformation, "Server Status"
End Sub

Private Sub wskserver_ConnectionRequest(ByVal requestID As Long)
If wskserver.State <> sckClosed Then wskserver.Close
wskserver.Accept requestID
End Sub

Private Sub wskserver_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim data As String
wskserver.GetData data
txtrecieve.Text = data
End Sub

Private Sub wskserver_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim data As String
wskserver.GetData data
txtrecieve.Text = data
End Sub
