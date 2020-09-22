VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5100
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "connection.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "connection.frx":27A2
   ScaleHeight     =   2550
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "&Join"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "&Host"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IDLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "REMOTE HOST:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then Text1.SetFocus: Exit Sub

Command3.Visible = True
Command1.Visible = False
Command2.Visible = False
Text1.Enabled = False
Text2.Enabled = False
lblStatus = "listening for an incomming connection .."
lblStatus.Visible = True


HostChat

End Sub

Private Sub Command2_Click()
If Text1 = "" Then Text1.SetFocus: Exit Sub
If Text2 = "" Then Text2.SetFocus: Exit Sub
Command3.Visible = True
Command1.Visible = False
Command2.Visible = False
Text1.Enabled = False
Text2.Enabled = False
lblStatus = "searching for chat session .."
lblStatus.Visible = True

JoinChat


End Sub

Private Sub Command3_Click()
Winsock1.Close
Winsock1_Close

End Sub

Private Sub Form_Load()
Text1 = GetSetting("one2onechat", "connection", "name")
Text2 = GetSetting("one2onechat", "connection", "remotehost")

End Sub

Private Sub Form_Terminate()
Winsock1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close

End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Form2, Button
End Sub

Private Sub Text1_Change()
SaveSetting "one2onechat", "connection", "name", Text1

End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)

End Sub

Private Sub Text2_Change()
SaveSetting "one2onechat", "connection", "remotehost", Text2
End Sub

Private Sub Text2_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub JoinChat()
Winsock1.Close
Winsock1.RemotePort = 45535
Winsock1.Connect Text2.Text, 45535

End Sub


Private Sub HostChat()
Winsock1.Close
Winsock1.LocalPort = 45535
Winsock1.Listen
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
lblStatus = "connection has been closed"
Command1.Visible = True
Command2.Visible = True
Text1.Enabled = True
Text2.Enabled = True
lblStatus.Visible = False
Command3.Visible = False

If Form1.Visible = True Then Form1.ConnectionClosed
End Sub

Private Sub Winsock1_Connect()
Winsock1.SendData "myname:" & Text1
Form2.Hide
Form1.GetName Text1
Form1.Show
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
Winsock1.SendData "myname:" & Text1.Text
Form2.Hide
Form1.GetName Text1
Form1.Show
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim incdata As String
Winsock1.GetData incdata
Form1.Arrival incdata

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
lblStatus = Description
'Command1.Visible = True
'Command2.Visible = True
'Text1.Enabled = True
'Text2.Enabled = True
'lblStatus.Visible = False
'Command3.Visible = False
If Form1.Visible = True Then Form1.ConnectionClosed

End Sub
