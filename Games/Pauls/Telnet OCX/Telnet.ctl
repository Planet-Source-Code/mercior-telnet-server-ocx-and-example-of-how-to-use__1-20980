VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Telnet 
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Telnet.ctx":0000
   ScaleHeight     =   3405
   ScaleWidth      =   4035
   ToolboxBitmap   =   "Telnet.ctx":0C42
   Begin MSWinsockLib.Winsock Ws 
      Index           =   0
      Left            =   0
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Telnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'[-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-]
'[ Telnet Server OCX by Paul Blower  ]
'[                                   ]
'[ Ive tried to comemnt all I can in ]
'[ the source for both OCX and the   ]
'[ example of how to use it.         ]
'[ ->>       Please Vote!        <<- ]
'[ Paul Blower : mercior@hotmail.com ]
'[-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-]

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim MaxCon As Integer
Dim UsePort As Integer
Dim TmpData() As String

Public Event DataIn(Index As Integer, Data As String)
Public Event Connection(Index As Integer, Port As Integer, RemIP As String)
Public Event Disconnect(Index As Integer)

Public Property Get ServerIP()
    ServerIP = Ws(0).LocalIP
End Property

Private Sub PS(Interval As Long) 'Pauses program on that line
    Start = GetTickCount
        Do While Start + Interval > GetTickCount
            DoEvents
        Loop
End Sub

Public Sub DisconnectUser(Index As Integer)
    Ws(Index).Close
    RaiseEvent Disconnect(Index)
End Sub

Public Sub CloseServer()
    For i = 0 To MaxCon 'close all sockets
        Ws(i).Close
    Next
End Sub

Public Sub SendToAll(Data As String)
    For i = 0 To MaxCon 'loop thru all sockets
        If Ws(i).State = sckConnected Then 'if the socket is connected
            Ws(i).SendData Data 'send the data
            PS 20 'pause to avoid data loss
        End If
    Next
End Sub

Public Sub StartServer(Port As Integer, MaxConnections As Integer)
    UsePort = Port
    MaxCon = MaxConnections
        For i = 1 To MaxCon 'Load all the nececarry sockets
            Load Ws(i)
        Next
    ReDim TmpData(MaxCon) As String
    Ws(0).LocalPort = Port
    Ws(0).Listen
End Sub

Public Sub SendData(Index As Integer, Data As String)
    Ws(Index).SendData Data
End Sub

Private Sub UserControl_Resize()
    Width = 32 * 15
    Height = 32 * 15
End Sub

Private Sub Ws_Close(Index As Integer) 'User has disconnected
    Ws(Index).Close
    RaiseEvent Disconnect(Index)
End Sub

Private Sub Ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Index <> 0 Then Exit Sub
    Ws(0).Close
    Ws(0).LocalPort = UsePort 'Keep our main socket listening on the server port
    Ws(0).Listen
        For i = 1 To MaxCon
            If Ws(i).State = sckClosed Then
                Ws(i).LocalPort = 5000 + i 'make a socket accept the next availiable socket
                Ws(i).Accept requestID     'Server starts from port 5000 and works its way up
                RaiseEvent Connection(CInt(i), 5000 + i, Ws(i).RemoteHostIP)
                Exit Sub
            End If
        Next
End Sub

Private Sub Ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
' For those who dont know, telnet sends data by sending each keystroke individually,
' which can make coding a server for it a bit of a pain. This algorithm below also
' allows for mixed data packets (eg if data isnt sent via each keystroke, but is mixed
' up and we get a packet thats a few characters long)
    Dim WData() As Byte
    ReDim WData(bytesTotal) As Byte
    Ws(Index).GetData WData
        For i = 0 To UBound(WData)
            If WData(i) = 8 Then 'Chr(8) is a telnet backspace... Telnet does horrible backspacing
                If Len(TmpData(Index)) > 0 Then
                    TmpData(Index) = Left(TmpData(Index), Len(TmpData(Index)) - 1)
                End If
            ElseIf WData(i) = 10 Or WData(i) = 13 Then 'Then user has pressed enter, so send the data theyd typed
                RaiseEvent DataIn(Index, TmpData(Index))
                TmpData(Index) = ""
                Exit For
            Else
                TmpData(Index) = TmpData(Index) & Chr(WData(i))
            End If
        Next
End Sub

Private Sub Ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Ws(Index).Close 'An error means a disconnection :)
    RaiseEvent Disconnect(Index)
End Sub
