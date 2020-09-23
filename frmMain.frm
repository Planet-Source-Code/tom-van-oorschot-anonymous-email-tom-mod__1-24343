VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anonymous E-Mailer"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNormal 
      Caption         =   "Normal"
      Height          =   225
      Left            =   540
      TabIndex        =   14
      Top             =   4290
      Width           =   1245
   End
   Begin VB.OptionButton optAnonymous 
      Caption         =   "Anonymous"
      Height          =   225
      Left            =   540
      TabIndex        =   13
      ToolTipText     =   "Will not always work, depends on servers configuration"
      Top             =   4560
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   4740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   780
      TabIndex        =   11
      Top             =   4860
      Width           =   1695
   End
   Begin VB.TextBox txtBody 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   2580
      Width           =   2895
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   2100
      Width           =   2295
   End
   Begin VB.TextBox txtFrom 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   1740
      Width           =   2295
   End
   Begin VB.TextBox txtTo 
      Appearance      =   0  'Flat
      Height          =   535
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Import from Textfile"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   12
      ToolTipText     =   "This method imports email adresses from a file that ere seperated with a comma (,)"
      Top             =   1410
      Width           =   1545
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   4920
      Width           =   540
   End
   Begin VB.Label lblSend 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2580
      TabIndex        =   9
      Top             =   4860
      Width           =   615
   End
   Begin VB.Label lblBody 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2370
      Width           =   420
   End
   Begin VB.Label lblSubj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   240
   End
   Begin VB.Line Line2 
      X1              =   630
      X2              =   630
      Y1              =   120
      Y2              =   780
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3360
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Shape Shape1 
      Height          =   5145
      Left            =   120
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Anonymous E-Mailer, coded by Patrick Moore, Modified by Tom van Oorschot"
      Height          =   585
      Left            =   720
      TabIndex        =   0
      Top             =   150
      Width           =   2400
   End
   Begin VB.Image img1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":0CCA
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************
'* CODE BY: ZELDA (PATRICK MOORE) *
'* MOD BY: "Tom van Oorschot      *
'* Credits go to Patrick Moore    *
'*                                *
'* Feel free to re-distribute or  *
'* Use in your own projects.      *
'* Giving credit to me would be   *
'* nice :)   -Patrick             *
'**********************************
'
'PS: Please look for more submissions to PSC by me
'    shortly.  I've recently been working on a lot
'    :))  All my submissions are under author name
'    "Patrick Moore"

Private Sub Form_Load()
 optNormal = True
End Sub

Private Sub Label1_Click()
Dim FileName As String, EmailData As String, LineData As String
Dim CurrentData As String, FindComma As String, TempData As String
Dim Counter As Integer
Counter = 1

'Shows a Commondialog ShowOpen box where the user can enter the path to the file
With CommonDialog1
  .Filter = "Text files (*.txt)|*.txt|DAT files (*.dat)|*.dat|All files (*.*)|*.*"
  .ShowOpen
  FileName = .FileName
End With

'Check if user didn't hit Cancel
If FileName = "" Then
  Exit Sub
End If
 
 'Open the file for input
 Open FileName For Input As #1
   'Loop trough the email address file
   Do Until EOF(1)
    Input #1, LineData
    EmailData = EmailData & LineData
   Loop
  Close #1
  
  'loop trough until last ; found
  Do Until FindComma = "0"
    If Counter = 1 Then
     FindComma = InStr(EmailData, ";")
    Else
     FindComma = InStr(CurrentData, ";")
    End If
    
      'Get the email address, this is the data before the found ; so the value of FindComma - 1 because we don't want the ;
     If Counter = 1 Then
      TempData = Left(EmailData, FindComma - 1)
     Else
      If FindComma <> "0" Then
       TempData = Left(CurrentData, FindComma - 1)
      Else
        TempData = CurrentData
      End If
     End If
   
   If Counter = 1 Then
      'Get the data behind the used email address
      CurrentData = Right(EmailData, Len(EmailData) - (Len(TempData) + 1))
    Else
     If FindComma <> "0" Then
      'Get the data behind the used email address
      CurrentData = Right(CurrentData, Len(CurrentData) - (Len(TempData) + 1))
    End If
   End If
     
     'Write email address to txtTo
     txtTo = txtTo & TempData & ","
     Counter = Counter + 1
   Loop
    'Delete last comma
    txtTo.Text = Left(txtTo.Text, Len(txtTo.Text) - 1)
End Sub

'Please note:
'The TO field..seperate entries must be
'seperated by commas, not crlf's

Private Sub lblSend_Click()
'Show the data form
frmWinsock.Show

'Setup mail server and port
frmWinsock.Winsock.RemoteHost = frmMain.txtServer.Text
frmWinsock.Winsock.RemotePort = 25

'Connect to mail server
frmWinsock.Winsock.Connect
End Sub
