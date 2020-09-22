VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
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
   Icon            =   "chatroom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "chatroom.frx":27A2
   ScaleHeight     =   3390
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   2160
   End
   Begin MSWinsockLib.Winsock BinSock 
      Left            =   4560
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   5055
   End
   Begin VB.Image FlashImage 
      Height          =   480
      Left            =   3960
      Picture         =   "chatroom.frx":3F684
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblFileStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "File Transfer Status"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "THE CHATROOM:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CHATROOM
Dim Name2 As String * 16
Dim MyName As String
Dim RemoteName As String
Dim NoButton As Boolean

Dim TheImage As Integer
Dim FlashedFor As Integer
Dim FlashFor As Integer

'BINARY SENDING
Dim BinSend As String
Dim BinSendFileName As String
Dim BinSendFileLen As String
Dim BinSendSoFar As Long

'BINARY RECEIVING
Dim BinRecFileName As String
Dim BinRecFileLen As String
Dim BinRec As String
Dim BinRecSoFar As Long

Dim BinaryInProgress As Boolean



Private Sub BinSock_Close()
lblFileStatus = "Binary connection closed"
BinSendFileName = ""
BinRecFileName = ""

BinRecSoFar = 0
BinSendSoFar = 0
BinRecFileLen = 0
BinSendFileLen = 0

BinaryInProgress = False
BinSock.Close
lblFileStatus.Visible = False

StartFlashing


End Sub

Private Sub BinSock_Connect()
lblFileStatus = "Binary connection established"
BinSock.SendData BinSend
End Sub

Private Sub BinSock_ConnectionRequest(ByVal requestID As Long)
BinSock.Close
BinSock.Accept requestID
lblFileStatus = "Waiting for " & BinRecFileName

End Sub

Private Sub BinSock_DataArrival(ByVal bytesTotal As Long)
BinSock.GetData BinData1$, vbByte

BinRec = BinRec & BinData1$
BinRecSoFar = BinRecSoFar + bytesTotal


PerCentage$ = Round(BinRecSoFar / BinRecFileLen * 100, 0) & "%"
lblFileStatus = "Downloading (" & PerCentage$ & ") .."


If BinRecSoFar = BinRecFileLen Then

Open App.Path & "\" & BinRecFileName For Binary As #1
Put #1, , BinRec
Close #1
lblFileStatus = "Download complete of " & BinRecFileName
StartFlashing
BinSock.Close
BinSock_Close
lblFileStatus.Visible = False

End If

End Sub

Private Sub BinSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lblFileStatus = "Binary connection error"
BinSendFileName = ""
BinRecFileName = ""

BinRecSoFar = 0
BinSendSoFar = 0
BinRecFileLen = 0
BinSendFileLen = 0

BinaryInProgress = False
BinSock.Close
End Sub

Private Sub BinSock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
On Error Resume Next

Dim PerCentage As String
BinSendSoFar = BinSendSoFar + bytesSent
PerCentage = BinSendSoFar / BinSendFileLength * 100
PerCentage = Round(BinSendSoFar / BinSendFileLen * 100, 0) & "%"
lblFileStatus = "Sending (" & PerCentage & ") .."
End Sub

Private Sub Command1_Click()
If NoButton = True Then Exit Sub
Chat MyName, Text2
SendChat Text2
Text2.SelStart = 0
Text2.SelLength = Len(Text2)
Text2.SetFocus


End Sub

Private Sub Chat(Name As String, ChatText As String)
Name2 = Name & " says: "
Text1 = Text1 & Name2 & " " & ChatText & vbCrLf
Text1.SelStart = Len(Text1)

End Sub

Private Sub Chat2(Text As String)
Text1 = Text1 & Text & vbCrLf
Text1.SelStart = Len(Text1)

End Sub

Private Sub SendChat(Text As String)
Form2.Winsock1.SendData "txt" & Text

End Sub

Private Sub Send(Text As String)
Form2.Winsock1.SendData Text
End Sub



Public Sub GetName(Name As String)
MyName = Name
End Sub
Public Sub Arrival(incdata As String)


If Mid(incdata, 1, 7) = "myname:" Then RemoteName = Mid(incdata, 8)

If Mid(incdata, 1, 3) = "txt" Then Chat RemoteName, Mid(incdata, 4): StartFlashing

If incdata = "filep" Then Chat2 " * " & RemoteName & " is preparing to send you a file": StartFlashing

If incdata = "filec" Then Chat2 " * " & RemoteName & " is no longer preparing to send you a file": lblFileStatus.Visible = False: BinSock.LocalPort = 0: BinSock.Close: StartFlashing

If Mid(incdata, 1, 9) = "filename:" Then
BinRecFileName = Mid(incdata, 10)
lblFileStatus = BinRecFileName
End If


If Mid(incdata, 1, 9) = "filesize:" Then
BinRecFileLen = Mid(incdata, 10)
Dim TheMessage2 As String
StartFlashing
TheMessage2 = "Do you want to accept " & BinRecFileLen & " bytes and download " & BinRecFileName
Debug.Print TheMessage2
response$ = MsgBox(TheMessage2, vbQuestion + vbYesNo, "File download")
If response$ = vbYes Then
BinaryInProgress = True
lblFileStatus = "Waiting for " & BinRecFileName
Send "accepted"
BinSock.Close
BinSock.LocalPort = 45533
BinSock.Listen
Else
Send "notaccepted"
End If
End If


If incdata = "accepted" Then
'CONNECT TO BINARY WINSOCK
BinSock.Close
BinSock.LocalPort = 0
BinSock.RemotePort = 45533
BinSock.Connect Form2.Winsock1.RemoteHostIP, 45533
End If

If incdata = "notaccepted" Then
MsgBox RemoteName & " did not accept your file", vbInformation, "File upload"
lblFileStatus.Visible = False
End If

End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
NoButton = True
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If BinaryInProgress = True Then
MsgBox "File transfer already in progress, cant send more than one file at any one time", vbExclamation, "Transfer in progress"
Exit Sub
End If


response$ = MsgBox("Would you like to send a file to " & RemoteName & "?", vbQuestion + vbYesNo, "Send file")
If response$ = vbNo Then NoButton = False: Exit Sub

Send "filep"
Dim BinSendFilename2 As String

BinSendFileName = FileOpen.GetFilename
BinSendFilename2 = BinSendFileName
If BinSendFileName = "cancel" Then Send "filec": NoButton = False: Exit Sub
lblFileStatus.Visible = True
lblFileStatus = "preparing a binary connection"


Open BinSendFileName For Binary As #1
BinSend = String(FileLen(BinSendFileName), " ")
Get #1, , BinSend
Close #1

Dim TempInt01 As Integer

BinSock.LocalPort = 45530
BinSock.RemotePort = 0

BinSock.Close


BinSendFileName = StrReverse(BinSendFileName)
TempInt01 = InStr(1, BinSendFileName, "\") - 1
BinSendFileName = Mid(BinSendFileName, 1, TempInt01)
BinSendFileName = StrReverse(BinSendFileName)

Send "filename:" & BinSendFileName
Sleep 200
BinSendFileLen = Len(BinSend)
Send "filesize:" & Len(BinSend)





NoButton = False
End Sub


Private Sub Form_GotFocus()
StopFlashing
End Sub

Private Sub Form_Load()
SysTray.AddTrayIcon Form1.Icon, Form1, "one2one chat - " & MyName
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TrayEvent As String
tray = SysTray.TrayEvent(X)

If tray = "LEFTUP" Then Form1.Show: StopFlashing

End Sub

Private Sub Form_Terminate()
SysTray.RemoveTrayIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
SysTray.RemoveTrayIcon

End Sub

Private Sub Label2_Click()
SysTray.RemoveTrayIcon
End

End Sub

Private Sub Label3_Click()
Form1.Visible = False

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Form1, Button

End Sub

Public Sub ConnectionClosed()
MsgBox RemoteName & " has left the conversation", vbInformation, "Disconnected"
SysTray.RemoveTrayIcon
End
End Sub

Private Sub lblFileStatus_Change()
lblFileStatus.Visible = True
End Sub

Private Sub Timer1_Timer()
If TheImage = 0 Then SysTray.ChangeTrayIcon FlashImage.Picture: TheImage = 1: GoTo 10
If TheImage = 1 Then SysTray.ChangeTrayIcon Form1.Icon: TheImage = 0: GoTo 10


10
If FlashFor = 0 Then Exit Sub
FlashedFor = FlashedFor + 1
If FlashedFor = FlashFor Then
StopFlashing
End If

End Sub

Private Sub StartFlashing(Optional HowLongFor As Integer = 0)

Timer1.Enabled = True
FlashFor = HowLongFor
If Form1.Visible = True Then FlashFor = 2

End Sub


Private Sub StopFlashing()
Timer1.Enabled = False
TheImage = 0
FlashedFor = 0
FlashFor = 0
SysTray.ChangeTrayIcon Form1.Icon
End Sub
