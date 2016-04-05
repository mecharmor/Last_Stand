'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsSound

    'Declares
    Private _strDirectory As String = ""
    Private _frmToPass As Form
    Private thrCloseFile As System.Threading.Thread
    Private intCounter As Integer = 0 'Used for alias
    Private intReturn As Integer = 0 'Check for errors
    Private strAlias As String = ""
    Private strRepeat As String = ""

    'Declare function
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String,
                                                                                   ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    Public Sub New(frmToPass As Form, strDirectory As String, intLengthOfFile As Integer, intVolume As Integer, Optional blnRepeat As Boolean = False)

        'Notes: Must pass a form and length of file, the length is the duration time, make sure to pass more than necessary, example: 00:02 seconds pass 3000 instead of 2000

        'Set
        intCounter = gintAlias + 1

        'Increase counter or start it over
        If gintAlias + 1 = Integer.MaxValue Then
            gintAlias = 0
        Else
            gintAlias += 1
        End If

        'Set
        _frmToPass = frmToPass

        'Set
        _strDirectory = strDirectory

        'Get alias
        strAlias = strGetAlias()

        'Check if repeating
        If blnRepeat Then
            'Set
            strRepeat = " repeat"
            'Prevent error
            Try
                'Open file
                _frmToPass.Invoke(Sub() intReturn = mciSendString("open " & ControlChars.Quote & _strDirectory & ControlChars.Quote & " alias " & strAlias, "", 0, 0))
                'Play file
                _frmToPass.Invoke(Sub() intReturn = mciSendString("play " & strAlias & strRepeat, "", 0, 0))
            Catch ex As Exception
                'No debug
            End Try
        Else
            'Set
            thrCloseFile = New System.Threading.Thread(New System.Threading.ThreadStart(Sub() ClosingFile(intLengthOfFile)))
            thrCloseFile.Start()
            'Play sound
            PlaySound()
        End If

        'Prevent error
        Try
            'Check for volume
            _frmToPass.Invoke(Sub() intReturn = mciSendString("setaudio " & strAlias & " volume to " & CStr(intVolume), "", 0, 0))
        Catch ex As Exception
            'No debug
        End Try

    End Sub

    Public Sub StopAndCloseSound()

        'Prevent error
        Try
            'Stop file
            _frmToPass.Invoke(Sub() intReturn = mciSendString("stop " & strAlias, "", 0, 0))
            'Close file
            _frmToPass.Invoke(Sub() intReturn = mciSendString("close " & strAlias, "", 0, 0))
        Catch ex As Exception
            'No debug
        End Try

    End Sub

    Private Function strGetAlias() As String

        'Notes: Looks like "X:\folder\folder\folder\file.mp3"

        'Declare
        Dim astrTemp1() As String = Split(_strDirectory, "\")
        Dim astrTemp2() As String = Split(astrTemp1(astrTemp1.GetUpperBound(0)), ".")

        'Return
        Return astrTemp2(0) & CStr(intCounter)

    End Function

    Private Sub PlaySound()

        'Notes: Must have a form because MCI uses a thread and needs an invoke, otherwise no sound

        'Prevent error
        Try
            'Open file
            _frmToPass.Invoke(Sub() intReturn = mciSendString("open " & ControlChars.Quote & _strDirectory & ControlChars.Quote & " alias " & strAlias, "", 0, 0))
            'Play file
            _frmToPass.Invoke(Sub() intReturn = mciSendString("play " & strAlias, "", 0, 0))
        Catch ex As Exception
            'No debug
        End Try

    End Sub

    Private Sub ClosingFile(intLengthOfFile As Integer)

        'Notes: Must have a form because MCI uses a thread and needs an invoke, otherwise no sound

        'Declare
        Dim intWait As Integer = 0

        'Sleep, but count so program can be safely exited immediately
        While intWait <> intLengthOfFile
            System.Threading.Thread.Sleep(1)
            intWait += 1
            If Not _frmToPass.IsHandleCreated Then
                Exit While
            End If
        End While

        'Prevent errors
        Try
            'Stop file
            _frmToPass.Invoke(Sub() intReturn = mciSendString("stop " & strAlias, "", 0, 0))
            'Close file
            _frmToPass.Invoke(Sub() intReturn = mciSendString("close " & strAlias, "", 0, 0))
        Catch ex As Exception
            'No debug
        End Try

    End Sub

End Class