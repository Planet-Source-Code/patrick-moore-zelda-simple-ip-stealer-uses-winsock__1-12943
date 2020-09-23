VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IP Stealer"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
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
   ScaleHeight     =   3855
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPatrick 
      Caption         =   "patrick's website"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "copy to clipboard"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Listen 
      Left            =   2400
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmMain.frx":08CA
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "80"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://0.0.0.0/"
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   3240
      Width           =   1080
   End
   Begin VB.Label lblTheURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "url:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   240
   End
   Begin VB.Line Line3 
      X1              =   3120
      X2              =   0
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " status "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   570
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   -360
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   0
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "listen on port:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label lblIP 
      BackStyle       =   0  'Transparent
      Caption         =   "0.0.0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   750
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblYourIP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "your IP:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************
'* CODE BY: PATRICK MOORE (ZELDA) *
'* Feel free to re-distribute or  *
'* Use in your own projects.      *
'* Giving credit to me would be   *
'* nice :)                        *
'*                                *
'* Please vote for me if you find *
'* this code useful :]   -Patrick *
'**********************************
'http://members.nbci.com/erx931/VB/
'
'PS: Please look for more submissions to PSC by me
'    I've recently been working on a lot of them.
'    :))  All my submissions are under author name
'    "Patrick Moore (Zelda)"

'Function related to the shell website
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub GotoURL()
Dim Success As Long, URL As String

URL = "http://members.nbci.com/erx931/VB/"
Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", 1)
End Sub

Sub Status(TheStatus As String)
'Send the status to the status textbox
txtStatus = txtStatus & TheStatus & vbCrLf
End Sub

Private Sub cmdCopy_Click()
'Set the clipboard to the URL
Clipboard.SetText lblURL.Caption
End Sub

Private Sub cmdListen_Click()
'Close the winsock
Listen.Close

'Set the local port to the one specified
Listen.LocalPort = txtPort.Text

'Listen for incoming connection requests
Listen.Listen

Status "STATUS> Listening.."
End Sub

Private Sub cmdPatrick_Click()
GotoURL
End Sub

Private Sub cmdRefresh_Click()
'Set the IP label's
'caption to your local IP
lblIP.Caption = Listen.LocalIP

'Set the URL label's caption to the URL
lblURL.Caption = "http://" & lblIP.Caption & "/"
End Sub

Private Sub cmdStop_Click()
'Close the winsock
Listen.Close
Status "STATUS> Listening haulted"
End Sub

Private Sub Form_Load()
'Set the IP label's
'caption to your local IP
lblIP.Caption = Listen.LocalIP

'Set the URL label's caption to the URL
lblURL.Caption = "http://" & lblIP.Caption & "/"
End Sub

Private Sub Listen_ConnectionRequest(ByVal requestID As Long)
'Close the winsock
Listen.Close

'Accept the connection request
Listen.Accept requestID

'Send the stolen IP to the textbox
Status "STOLEN> " & Listen.RemoteHostIP
End Sub


Private Sub Listen_DataArrival(ByVal bytesTotal As Long)
Dim HTML As String
'Data has come in

'Edit the following line with the website
'you would like the client to see
HTML = "<HTML><HEAD><TITLE>404 Error: File Not Found</TITLE></HEAD><BODY><FONT FACE=""TAHOMA"" SIZE=""-1""><B>404</B> ERROR:<BR>File Not Found<BR><BR>The specified file was not found on the server.</BODY></HTML>"

'Send a false webpage to the client
Listen.SendData "HTTP/1.1 200 OK" & vbCrLf & _
"Server: IP32/0.0.1 (Win32)" & vbCrLf & _
"Accept -Ranges: bytes" & vbCrLf & _
"Content-Length: " & Len(HTML) & vbCrLf & _
"Connection: close" & vbCrLf & _
"Content-Type: text/html" & vbCrLf & vbCrLf & _
HTML & vbCrLf

End Sub

Private Sub Listen_SendComplete()
'The false webpage has been sent to
'the client, now we must..

'Close the winsock control
Listen.Close

'Set the local port to the one specified
Listen.LocalPort = txtPort.Text

'Await new incoming connection requests
Listen.Listen
End Sub

Private Sub txtStatus_Change()
'Set the cursor to the last character
'of the textbox
txtStatus.SelStart = Len(txtStatus)
End Sub
