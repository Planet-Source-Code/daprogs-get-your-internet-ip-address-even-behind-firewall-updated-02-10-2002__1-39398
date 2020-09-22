VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Caption         =   "Get your current Internet IP"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAddr 
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      Top             =   360
      Width           =   2835
   End
   Begin VB.TextBox txtIP 
      Height          =   255
      Left            =   1260
      TabIndex        =   4
      Top             =   60
      Width           =   2835
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Get Inet Info"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   60
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Host:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current IP:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Entire code by Didier Aeschimann
'Copyright Didier Aeschimann
'Please do not use without permission of author

'This will use web server on internet to get you real ip address
'even if you are behind a firewall or proxy

Dim strData As String
Dim strIP As String
Dim strAddr As String

Private Sub cmdExecute_Click()

'Get data from web server
strData = Inet1.OpenURL("www.daprogs.com/ip/?uid=psc021002")

'Store IP to variable
strIP = ParseXML(strData, "IP")
'Store Host to variable
strAddr = ParseXML(strData, "ADDR")

'Display values
txtIP.Text = strIP
txtAddr.Text = strAddr

End Sub

Private Sub cmdQuit_Click()
  End
End Sub
