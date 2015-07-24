VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck_1.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "OtadChat: Client Version"
   ClientHeight    =   7155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6135
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5880
         Top             =   1320
      End
      Begin VB.CommandButton cmddisconnect 
         Caption         =   "Logout"
         Height          =   495
         Left            =   4800
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdconnect 
         Caption         =   "Login"
         Height          =   495
         Left            =   4800
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtname 
         Height          =   405
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtip 
         Height          =   405
         Left            =   1560
         TabIndex        =   7
         Text            =   "localhost"
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1035
      End
   End
   Begin VB.Frame chat 
      BackColor       =   &H00FFC0C0&
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   6015
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   6240
         TabIndex        =   13
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   4560
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Client Form.frx":0000
         Left            =   0
         List            =   "Client Form.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4560
         Width           =   3015
      End
      Begin VB.TextBox txtoutput 
         Height          =   4095
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   480
         Width           =   6015
      End
      Begin VB.CommandButton cmdsend 
         Caption         =   "Send"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox txtsend 
         Height          =   765
         Left            =   0
         TabIndex        =   1
         Top             =   4920
         Width           =   6015
      End
   End
   Begin MSWinsockLib.Winsock tcpclient 
      Left            =   1320
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public clientname As String
Public hostname As String
Public errorcount As Integer
Dim exists As String
Private Sub about_Click()
Call Load(Form2)
Form2.Show
End Sub

Private Sub Close_Click()
End
End Sub

Private Sub cmdclear_Click()
txtoutput.Text = ""
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus

End Sub

Private Sub cmdconnect_Click()
If txtip.Text <> "" And txtname.Text <> "" Then
cmdconnect.Enabled = False
Dim i As Integer
txtoutput.FontSize = 9
clientname = txtname.Text
tcpclient.Close
tcpclient.remoteport = remoteport
tcpclient.RemoteHost = txtip.Text
tcpclient.Connect

Else
Dim r As Integer
r = MsgBox("Enter Client Name!", vbOKOnly, "Error")
End If

End Sub

Private Sub cmddisconnect_Click()
cmddisconnect.Enabled = False
cmdconnect.Enabled = True
Call tcpclient.Close
txtoutput.Text = ""
txtoutput.Text = "Disconnected from Server"
Combo1.Clear
chat.Enabled = False
End Sub

Private Sub cmdsend_Click()
If Combo1.Text <> "" Then
Call tcpclient.SendData("\message\" & Combo1.Text & "\" & clientname & " >>>" & " " & txtsend.Text)
txtoutput.Text = txtoutput.Text & vbCrLf & clientname & " >>>" & " " & txtsend.Text & vbCrLf
exists = clientname & " >>>" & " " & txtsend.Text
txtsend.Text = ""
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
End If
End Sub



Private Sub Form_Load()
Timer1.Interval = 10
Timer1.Enabled = False
chat.Enabled = False
cmddisconnect.Enabled = False
remoteport = 5000

End Sub

Private Sub Form_Terminate()
Call tcpclient.Close
End Sub

Private Sub tcpclient_Close()
cmddisconnect.Enabled = False
Call tcpclient.Close
txtoutput.Text = "Server Closed Connection"
cmdconnect.Enabled = True
Combo1.Clear
txtoutput.SelStart = Len(txtoutput.Text)
cmdconnect.SetFocus
hostname = ""
chat.Enabled = False
End Sub

Private Sub tcpclient_Connect()
cmdconnect.Enabled = False
cmddisconnect.Enabled = True
chat.Enabled = True
txtoutput.Text = "Connected to IP Address: " & tcpclient.RemoteHostIP & vbCrLf _
& "Port #: " & tcpclient.remoteport & vbCrLf
Dim name As String
name = "\newclientname\" & clientname
Call tcpclient.SendData(name)
Combo1.AddItem ("Everyone")
errorcount = 0
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
End Sub

Private Sub tcpclient_DataArrival(ByVal bytesTotal As Long)
Dim message As String
Call tcpclient.GetData(message)

'''''''''''''''''''''''''''''''''''''''
Dim nmessage As String
nmessage = message
Dim name As String
Dim pos As Integer
pos = InStr(message, "\servername\")
If pos = 1 Then
name = Right$(message, (Len(message) - Len("\servername\")))
Combo1.AddItem (name)
hostname = name

Exit Sub
End If
''''''''''''''''''''''''''''''''''''''
Dim mmessage As String
mmessage = message
Dim pos2 As Integer
pos2 = InStr(mmessage, "\message\")
If pos2 = 1 Then
mmessage = Right$(mmessage, (Len(mmessage) - Len("\message\")))

If mmessage = exists Then
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
Exit Sub
End If
txtoutput = txtoutput & vbCrLf & mmessage & vbCrLf
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''
Dim lmessage As String
lmessage = message
Dim pos3 As Integer
pos3 = InStr(lmessage, "\serverlist\")
If pos3 = 1 Then
Call updatelist(lmessage)
End If
End Sub


Private Sub tcpclient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Number = 10061 Then
If errorcount < 20 Then
tcpclient.Close
remoteport = remoteport + 1
tcpclient.remoteport = remoteport
tcpclient.RemoteHost = txtip.Text
errorcount = errorcount + 1
Call connection
Exit Sub
Else
MsgBox "Failed to connect to Server 20 Times! Server Full!", vbSystemModal
cmddisconnect.Enabled = False
cmdconnect.Enabled = True
Call tcpclient.Close
remoteport = 5000
txtoutput.Text = "Failed to connect to server"
Combo1.Clear
Exit Sub
End If
End If
Dim result As Integer
result = MsgBox(Source & ":  " & Description, vbOKOnly, "TCP/IP ERROR")
MsgBox Number
End Sub

Private Sub connection()
Call tcpclient.Connect
End Sub


Private Sub updatelist(message As String)
Dim result As Integer
Call Combo1.Clear
Call Combo1.AddItem("Everyone")
Call Combo1.AddItem(hostname)
Combo1.Text = hostname

message = Right$(message, (Len(message) - 12))
Dim pos As Integer
Dim entry As String
Do While Len(message) > 0
pos = InStr(message, "\")
entry = Left$(message, (pos - 1))
result = StrComp(entry, clientname)
If result <> 0 Then
Call Combo1.AddItem(entry)
End If
If pos = Len(message) Then
Exit Sub
End If
message = Right$(message, ((Len(message)) - pos))
Loop
End Sub


Private Sub txtsend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Combo1.Text <> "" Then
Call tcpclient.SendData("\message\" & Combo1.Text & "\" & clientname & " >>>" & " " & txtsend.Text)
txtoutput.Text = txtoutput.Text & vbCrLf & clientname & " >>>" & " " & txtsend.Text & vbCrLf
exists = clientname & " >>>" & " " & txtsend.Text
txtsend.Text = ""
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
End If
End If
End Sub
