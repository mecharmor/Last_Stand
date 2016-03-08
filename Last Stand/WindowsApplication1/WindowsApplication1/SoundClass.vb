'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class SoundClass

    'Declare
    Private _strName As String = ""
    Private _strDirectory As String = ""
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String,
                                                                                   ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    Public Sub StopSound()

        'Stop sound
        mciSendString("stop " & _strName, "", 0, 0)

    End Sub

    Public Sub New(strName As String, strDirectory As String)

        'Set
        _strName = strName
        _strDirectory = strDirectory

    End Sub

    Public Sub PlaySound(blnRepeat As Boolean, Optional intVolume As Integer = 1000)

        'Check directory
        If IO.File.Exists(_strDirectory) Then

            'Close the play if it was played
            Try
                mciSendString("close " & _strName, "", 0, 0)
            Catch ex As Exception
                'No debug
            End Try

            'Declare
            Dim strRepeat As String = ""

            'Check for repeat
            If blnRepeat Then
                strRepeat = " repeat"
            End If

            'Prepare file for open
            mciSendString("open " & Chr(34) & _strDirectory & Chr(34) & " alias " & _strName, "", 0, 0)

            'Play
            mciSendString("play " & _strName & strRepeat, "", 0, 0)

            'Optional change volume
            mciSendString("setaudio " & _strName & " vole to " & CStr(intVolume), "", 0, 0)

        End If

    End Sub

End Class
