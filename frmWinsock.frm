VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWinsock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Incoming / Outgoing Data"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWinsock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   2040
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   25
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'placed the buffer var in option explicit so it can be used by other subs
Dim Buffer As String
Sub WinsockReady()
'Execute until winsock receives data
Do
  DoEvents
Loop Until Buffer <> ""
End Sub
Sub Pause(duration)
'Pause for the specified duration
'Duration is in seconds
Dim Current As Long
Current = Timer
Do Until Timer - Current >= duration
    DoEvents
Loop
End Sub

Sub Status(data As String)
'Update the Data textbox
txtData.Text = txtData.Text & data & vbCrLf
End Sub

Sub SendData(data As String)
'Send this data to the server
Winsock.SendData data & vbCrLf

'Display what we're sending to the server
Status "> " & data
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Winsock.Close
Unload Me
End Sub

Private Sub txtData_Change()
'Automatically scroll to the last line
txtData.SelStart = Len(txtData.Text)
End Sub

Private Sub Winsock_Close()
'Winsock closed
Status "<CONNECTION CLOSED>"
End Sub

Private Sub Winsock_Connect()
Dim SendTo As String, SentFrom As String, Subject As String, MailBody As String
SendTo = frmMain.txtTo.Text
SentFrom = frmMain.txtFrom.Text
Subject = frmMain.txtSubject.Text
MailBody = frmMain.txtBody.Text

'Connected to server
Status "<CONNECTED>"

'Begin sending mail data

'Identify to the mail server

SendData "HELO x.x"
Call WinsockReady
'Pause 0.3

'Identify who the mail is from
If frmMain.optNormal = True Then
  SendData "MAIL FROM: " & SentFrom
Else
  SentFrom = InputBox("What (fake) email address you want to use?", "Fake Email address")
  If SentFrom <> "" Then
    SendData "MAIL FROM: " & SentFrom
  Else
    Exit Sub
  End If
End If

'Send to all recipients in the To field
If Right(SendTo, 1) <> "," Then SendTo = SendTo & ","
Do
    'Pause 0.5
    Call WinsockReady
    'Identify one recipient
    SendData "RCPT TO: " & Mid(SendTo, 1, InStr(SendTo, ",") - 1)
    SendTo = Mid(SendTo, InStr(SendTo, ",") + 1, Len(SendTo))
    
    'Continue trimming the recipient string
Loop Until InStr(SendTo, ",") = 0

Pause 0.5

'Tell the server we're ready to send the mail body and subject
SendData "DATA"
Call WinsockReady
'Pause 0.9

'Send the subject
SendData "SUBJECT: " & Subject
Call WinsockReady
'Pause 0.3

'Send the mail body
SendData MailBody

'End the mail body
SendData vbCrLf & vbCrLf & "."

'Wait for server to catch up
Call WinsockReady
'Pause 3

'Close connection to the server
Winsock.Close
Winsock_Close
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)

'Get incoming data, set that data to the Buffer variable
Winsock.GetData Buffer

'Send the data to the Data textbox
Status Buffer
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Error ocurred, display in Data textbox
Status "<WINSOCK ERROR> " & Number & ". " & Description & "."
End Sub
