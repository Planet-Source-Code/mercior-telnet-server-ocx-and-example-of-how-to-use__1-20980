VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "*\A..\..\TELNET~1\TELNET~1\Project1.vbp"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telnet OCX Example"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin TelnetControl.Telnet Telnet 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Server"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Boot Selected User"
      Height          =   315
      Left            =   6960
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Send Data to Selected Socket Only"
      Height          =   615
      Left            =   6960
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ListBox UList 
      Height          =   2010
      ItemData        =   "Form1.frx":0000
      Left            =   6960
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5895
   End
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   285
      Left            =   6090
      TabIndex        =   0
      Top             =   3720
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    If UList.ListIndex >= 0 Then 'Make sure theres a selection
        If Not Left(UList.List(UList.ListIndex), 6) = "Socket" Then 'Check the socket has a conenction
            Telnet.DisconnectUser UList.ListIndex + 1 'disconnect the user
        End If
    End If
End Sub

Private Sub Command2_Click()
    If Check1.Value = 0 Then 'Sends what was in the textbox to everyone in the server
        Telnet.SendToAll "Server: " & Text1.Text & vbCrLf
    Else
        If UList.ListIndex >= 0 Then 'Make sure theres a selection
            If Not Left(UList.List(UList.ListIndex), 6) = "Socket" Then 'Check the socket has a conenction
                Telnet.SendData UList.ListIndex + 1, "Server: " & Text1.Text & vbCrLf
            End If
        End If
    End If
    Text1.Text = ""
End Sub

Private Sub Command3_Click()
    Telnet.CloseServer 'well, im sure you can guess this one :P
End Sub

Private Sub Form_Load()
    Telnet.StartServer 23, 10 'Start the telnet server on port 23 and allow 10 connections max

    Text2.Text = "IP: " & Telnet.ServerIP 'set the textbox to the servers ip (your ip)
    
    For i = 0 To 9
        UList.AddItem "Socket " & i, i 'Fill up the socket list
    Next
End Sub

Private Sub Telnet_Connection(Index As Integer, Port As Integer, RemIP As String)
    'when somebody connects to the server, this sub is called
    SetUserList Index - 1, Index & ") " & RemIP  'fill in the socket box
End Sub

Sub SetUserList(LNum As Integer, NewText As String) 'just a sub that writes an item to the listbox
    UList.RemoveItem LNum
    UList.AddItem NewText, LNum
End Sub

Private Sub Telnet_DataIn(Index As Integer, Data As String)
    'When you press enter in telnet, the data is sent to this sub
    LogTxt "User " & Index & ": " & Data
End Sub

Private Sub Telnet_Disconnect(Index As Integer)
    'If somebody disconnects from the server this sub is called
    SetUserList Index - 1, "Socket " & Index - 1 'fill in the socket list
End Sub

Sub LogTxt(Txt As String) 'just a sub for writing to the richtextbox
    Rtb.Text = Rtb.Text & Txt & vbCrLf
    Rtb.SelStart = Len(Rtb.Text)
End Sub
