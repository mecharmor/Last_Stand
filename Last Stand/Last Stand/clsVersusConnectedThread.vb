'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsVersusConnectedThread

    'Declare
    Private blnReadLine As Boolean = False
    Private thrReadLine As System.Threading.Thread
    Private blnCheckConnection As Boolean = False
    Private thrCheckConnection As System.Threading.Thread
    Private blnConnectionLost As Boolean = False

    'Events used for tcp/ip connections
    Public Event gGotData(strData As String)
    Public Event gConnectionGone()

    Public Sub New(tcpcClient As System.Net.Sockets.TcpClient, srData As System.IO.StreamReader)

        'Start a read line thread
        blnReadLine = True
        thrReadLine = New System.Threading.Thread(Sub() ReadLine(srData))
        thrReadLine.Start()

        'Start a check connection thread
        blnCheckConnection = True
        thrCheckConnection = New System.Threading.Thread(Sub() CheckConnection(tcpcClient))
        thrCheckConnection.Start()

    End Sub

    Private Sub ReadLine(srData As System.IO.StreamReader)

        'Declare
        Dim strData As String = ""

        'Loop
        While blnReadLine
            'Get data
            Try
                strData = srData.ReadLine 'Blocks, gets stuck here, this is why it is in a thread
            Catch ex As Exception
                'Debug
                'Debug.Print(ex.ToString)
                'Exit
                blnConnectionLost = True
                Exit Sub
            End Try
            'Raise event, data ready
            RaiseEvent gGotData(strData)
        End While

    End Sub

    Private Sub CheckConnection(tcpcClient As System.Net.Sockets.TcpClient)

        'Loop
        While blnCheckConnection
            'Reduce CPU usage
            System.Threading.Thread.Sleep(1) 'If this line doesn't exist, a problem occurs with MCIsendstring, keep this line, strange error between threads
            'Check connection
            If Not blnStillConnected(tcpcClient) Or blnConnectionLost Then
                'Raise event
                RaiseEvent gConnectionGone()
                'Exit
                blnReadLine = False
                Exit Sub
            End If
        End While

    End Sub

    Private Function blnStillConnected(tcpcClient As System.Net.Sockets.TcpClient) As Boolean

        'Check connection
        Try
            If tcpcClient.Client.Poll(0, System.Net.Sockets.SelectMode.SelectRead) Then
                'Declare
                Dim bytBuffer(1) As Byte
                'Check state
                If tcpcClient.Client.Receive(bytBuffer, System.Net.Sockets.SocketFlags.Peek) = 0 Then
                    'Return, connection is lost
                    Return False
                End If
            End If
        Catch ex As Exception
            'Debug
            Return False
        End Try

        'Return, connection is good
        Return True

    End Function

    Public Sub AbortReadLineThread()

        'Abort
        Try
            blnReadLine = False
            thrReadLine.Abort()
        Catch ex As Exception
            'No debug
        End Try

    End Sub

    Public Sub AbortCheckConnectionThread()

        'Abort
        Try
            blnCheckConnection = False
            thrCheckConnection.Abort()
        Catch ex As Exception
            'No debug
        End Try

    End Sub

End Class