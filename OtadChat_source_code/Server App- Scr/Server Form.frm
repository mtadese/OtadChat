VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck_1.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "OtadChat: Server Version"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer send 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   5160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtserverlimit 
         Height          =   405
         Left            =   1440
         TabIndex        =   11
         Text            =   "50"
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmddisconnect 
         Caption         =   "Disconnect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdhost 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtname 
         Height          =   405
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label hostip 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Limit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame chat 
      BackColor       =   &H00800000&
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   5775
      Begin VB.CommandButton cmdkickclient 
         Caption         =   "Remove User"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtsend 
         Height          =   765
         Left            =   0
         TabIndex        =   4
         Top             =   4800
         Width           =   5775
      End
      Begin VB.CommandButton cmdsend 
         Caption         =   "Send"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtoutput 
         Height          =   4455
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   0
         Width           =   5775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4440
         Width           =   1695
      End
   End
   Begin MSWinsockLib.Winsock tcpclient 
      Left            =   1200
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpserver 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hostname As String
Dim nameindex As New Dictionary
Dim indexindex As New Dictionary

Dim sendindex1 As New Dictionary
Dim sendindex2 As New Dictionary

Dim connection As Boolean
Dim exists As Integer
Dim current As String




Private Sub About_Click()
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

Private Sub cmddisconnect_Click()
cmddisconnect.Enabled = False
Call nameindex.RemoveAll
Call indexindex.RemoveAll

Call sendindex1.RemoveAll
Call sendindex2.RemoveAll
send.Enabled = False
If connection = True Then
Dim X As Integer
For X = 0 To Val(txtserverlimit.Text - 1)
Call tcpserver(X).Close
If X <> 0 Then
Unload tcpserver(X)
End If
Next X
End If
cmdhost.Enabled = True
txtoutput.Text = "Closed all Connections!"
txtsend.Text = ""
cmdhost.SetFocus
End Sub

Private Sub cmdhost_Click()
If txtname.Text <> "" Then
send.Enabled = True
hostname = txtname.Text
Dim localport As Long
localport = 5000
Dim X As Integer
For X = 0 To Val(txtserverlimit.Text - 1)
If X <> 0 Then
Load tcpserver(X)
End If
tcpserver(X).localport = localport
tcpserver(X).Listen
localport = localport + 1
Next X

cmdhost.Enabled = False
cmddisconnect.Enabled = True
connection = True
txtoutput.FontSize = 9
txtoutput.Text = "Waiting for Connections" & vbCrLf
exists = 2
hostip.Caption = "Server IP Address: " & tcpserver(0).LocalIP
Else
Dim r As Integer
r = MsgBox("Enter Host Name", vbOKOnly, "Error")
End If
End Sub


Private Sub cmdkickclient_Click()
If Combo1.Text = "Everyone" Then
Exit Sub
End If

Dim index As Integer
index = nameindex.Item(Combo1.Text)
Call tcpserver(index).Close
Call tcpserver(index).Listen
Dim name As String
name = indexindex.Item(index)
Call indexindex.Remove(index)
Call nameindex.Remove(name)
txtoutput = txtoutput.Text & vbCrLf & vbCrLf & name & " was kicked from server!" & vbCrLf
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
Call serverlistupdate

Dim i As Integer
Dim value As Integer
Dim list() As Variant
Dim message As String
Dim message2 As String
message = "\message\" & (name & " was kicked from server!")
list = indexindex.Keys()
For i = 0 To indexindex.Count - 1
message2 = message

Do While sendindex1.exists(value) = True
value = value + 1
Loop
Call sendindex1.Add(value, list(i))
Call sendindex2.Add(value, message2)
'''''''''''''''''''''''''''''''''''''''
Next i

End Sub

Private Sub cmdsend_Click()
If Combo1.Text <> "" Then
Dim message As String
Dim name As String
message = hostname & " >>>" & " " & txtsend.Text
name = Combo1.Text
txtoutput.Text = txtoutput.Text & vbCrLf & message & vbCrLf
txtsend.Text = ""
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
'''''''''''''''''''''''''
If name = "Everyone" Then
Dim value As Integer
Dim i As Integer
Dim list() As Variant
Dim message2 As String
list = indexindex.Keys()
For i = 0 To indexindex.Count - 1
message2 = message
message2 = "\message\" & message2

Do While sendindex1.exists(value) = True
value = value + 1
Loop
Call sendindex1.Add(value, list(i))
Call sendindex2.Add(value, message2)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Next i
'''''''''''''''''''''''''''
Else
Dim message3 As String
message3 = message
message3 = "\message\" & message3
Dim value2 As Integer
value2 = nameindex.Item(name)
Call tcpserver(value2).SendData(message3)
End If
''''''''''''''''''''''''''''


End If
End Sub





Private Sub Combo1_Change()
current = Combo1.Text
End Sub

Private Sub Form_Load()
chat.Enabled = False
cmddisconnect.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
If connection = True Then
Dim X As Integer
For X = 0 To Val(txtserverlimit.Text - 1)
Call tcpserver(X).Close
If X <> 0 Then
Unload tcpserver(X)
End If
Next X
End If

End Sub


Private Sub serverlistupdate()
current = Combo1.Text
Dim i As Integer
Dim X As Integer
Dim elements() As Variant
Dim list() As Variant
elements = nameindex.Keys()
list = indexindex.Keys()
Dim index As Integer
Dim message As String
message = message & "\serverlist\"
For X = 0 To indexindex.Count - 1
index = list(X)
For i = 0 To nameindex.Count - 1
message = message & elements(i) & "\"
Next i
Dim value As Integer
Do While sendindex1.exists(value) = True
value = value + 1
Loop
Call sendindex1.Add(value, index)
Call sendindex2.Add(value, message)


message = ""
message = message & "\serverlist\" '
Next X
Dim e() As Variant
e = nameindex.Keys()
Dim b As Integer
Combo1.Clear
Combo1.AddItem "Everyone"
For b = 0 To nameindex.Count - 1
Combo1.AddItem (e(b))
Next b

If nameindex.exists(current) = True Then
Combo1.Text = current
Else
Combo1.Text = "Everyone"
End If

End Sub

Private Sub send_Timer()
send.Enabled = False
If sendindex1.Count = 0 Then
send.Enabled = True
Exit Sub
End If
Dim elements() As Variant
elements = sendindex1.Keys()
Dim index As Integer
Dim message As String
index = sendindex1.Item(elements(0))
message = sendindex2.Item(elements(0))
If tcpserver(index).State = sckConnected Then
Call tcpserver(index).SendData(message)
End If
Call sendindex1.Remove(elements(0))
Call sendindex2.Remove(elements(0))
send.Enabled = True
End Sub

Private Sub tcpserver_Close(index As Integer)
Call tcpserver(index).Close
Call tcpserver(index).Listen
Dim name As String
name = indexindex.Item(index)
Call indexindex.Remove(index)
Call nameindex.Remove(name)
txtoutput = txtoutput.Text & vbCrLf & name & " disconnected from server!" & vbCrLf
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
Call serverlistupdate

Dim i As Integer
Dim value As Integer
Dim list() As Variant
Dim message As String
Dim message2 As String
message = "\message\" & (name & " disconnected from server!")
list = indexindex.Keys()
For i = 0 To indexindex.Count - 1
message2 = message
Do While sendindex1.exists(value) = True
value = value + 1
Loop
Call sendindex1.Add(value, list(i))
Call sendindex2.Add(value, message2)


Next i


End Sub

Private Sub tcpserver_ConnectionRequest(index As Integer, ByVal requestID As Long)
If tcpserver(index).State <> sckClosed Then
Call tcpserver(index).Close
End If
chat.Enabled = True
Call tcpserver(index).Accept(requestID)
txtoutput.Text = txtoutput.Text & vbCrLf & "Connection from IP Address: " & tcpserver(index).RemoteHostIP & vbCrLf & "Port #: " & tcpserver(index).RemotePort & vbCrLf
Call tcpserver(index).SendData("\servername\" & hostname)
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
End Sub

Private Sub tcpserver_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim message As String
Call tcpserver(index).GetData(message)

'''''''''''''''''''''''''''''''''''''''
Dim nmessage As String
nmessage = message

Dim pos As Integer
pos = InStr(nmessage, "\newclientname\")
If pos = 1 Then
Call addname(nmessage, index)
Call serverlistupdate

Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''
Dim mmessage As String
mmessage = message
Dim pos2 As Integer
pos2 = InStr(mmessage, "\message\")
If pos2 = 1 Then
Call processmessage(mmessage)
Exit Sub
End If
End Sub

Private Sub tcpserver_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim result As Integer
result = MsgBox(Source & ":  " & Description, vbOKOnly, "TCP/IP ERROR")
End Sub

Public Sub addname(nmessage As String, index As Integer)

nmessage = Right$(nmessage, (Len(nmessage) - Len("\newclientname\")))
If nameindex.exists(nmessage) Then
Call tcpserver(index).Close
Call tcpserver(index).Listen
Else
Call nameindex.Add(nmessage, index)
Call indexindex.Add(index, nmessage)
Combo1.AddItem (nmessage)

Dim i As Integer
Dim list() As Variant
Dim Text1 As String
Dim text2 As String
Dim value As Integer
Text1 = ("\message\" & nmessage & " joined server!")
list = indexindex.Keys()
For i = 0 To indexindex.Count - 1
text2 = Text1
If (list(i)) <> index Then

Do While sendindex1.exists(value) = True
value = value + 1
Loop
Call sendindex1.Add(value, list(i))
Call sendindex2.Add(value, text2)
''''''''''''''''''''''''''
End If
Next i

End If
End Sub

Public Sub processmessage(message As String)
message = Right$(message, Len(message) - 9)
Dim name As String
Dim pos As Integer
Dim index As Integer
pos = InStr(message, "\")
name = Left$(message, (pos - 1))
message = Right$(message, Len(message) - pos)
'''''''''''''''''''''''''''''''''''''''''''''''''
If name = hostname Then
txtoutput.Text = txtoutput.Text & vbCrLf & message & vbCrLf
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
If name = "Everyone" Then
txtoutput.Text = txtoutput.Text & vbCrLf & message & vbCrLf
Dim i As Integer
Dim list() As Variant
Dim value As Integer
Dim message2 As String
list = indexindex.Keys()
For i = 0 To indexindex.Count - 1
message2 = message
message2 = "\message\" & message2

Do While sendindex1.exists(value) = True
value = value + 1
Loop
Call sendindex1.Add(value, list(i))
Call sendindex2.Add(value, message2)

Next i
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
index = nameindex.Item(name)
message = "\message\" & message
Call tcpserver(index).SendData(message)
End Sub


Private Sub txtsend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Combo1.Text <> "" Then
Dim message As String
Dim name As String
message = hostname & " >>>" & " " & txtsend.Text
name = Combo1.Text
txtoutput.Text = txtoutput.Text & vbCrLf & message & vbCrLf
txtsend.Text = ""
txtoutput.SelStart = Len(txtoutput.Text)
txtsend.SetFocus
'''''''''''''''''''''''''
If name = "Everyone" Then
Dim value As Integer
Dim i As Integer
Dim list() As Variant
Dim message2 As String
list = indexindex.Keys()
For i = 0 To indexindex.Count - 1
message2 = message
message2 = "\message\" & message2

Do While sendindex1.exists(value) = True
value = value + 1
Loop
Call sendindex1.Add(value, list(i))
Call sendindex2.Add(value, message2)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Next i
'''''''''''''''''''''''''''
Else
Dim message3 As String
message3 = message
message3 = "\message\" & message3
Dim value2 As Integer
value2 = nameindex.Item(name)
Call tcpserver(value2).SendData(message3)
End If
''''''''''''''''''''''''''''


End If
End If

End Sub

