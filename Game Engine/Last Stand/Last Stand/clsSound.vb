﻿'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsSound

    'Declares
    Private _strDirectory As String = ""
    Private _frmToPass As Form
    Private sttTimer As System.Timers.Timer
    Private intCounter As Integer = 0 'Used for alias
    Private intReturn As Integer = 0 'Check for errors
    Private strAlias As String = ""
    Private strRepeat As String = ""

    'Declare function
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String,
                                                                                   ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    'Declare delegates
    Private Delegate Sub PlaySoundDelegate()
    Private Delegate Sub ClosingSoundDelegate(sender As Object, e As System.Timers.ElapsedEventArgs)

    Public Sub New(frmToPass As Form, strDirectory As String, intLengthOfFile As Integer, intVolume As Integer, Optional blnRepeat As Boolean = False)

        'Notes: Must pass a form and length of file, the length is the duration time, make sure to pass more than necessary, example: 00:02 seconds pass 3000 instead of 2000

        'Set
        intCounter = gintAlias + 1

        'Increase counter or start it over
        If gintAlias + 1 >= 50000 Then
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
            'Open file
            intReturn = mciSendString("open " & ControlChars.Quote & _strDirectory & ControlChars.Quote & " alias " & strAlias, "", 0, 0)
            'Play file
            intReturn = mciSendString("play " & strAlias & strRepeat, "", 0, 0)
        Else
            'Set
            sttTimer = New System.Timers.Timer(intLengthOfFile)
            AddHandler sttTimer.Elapsed, AddressOf ClosingFile
            sttTimer.AutoReset = False
            'Play sound
            PlaySound()
            'Start timer
            sttTimer.Enabled = True
        End If

        'Check for volume
        intReturn = mciSendString("setaudio " & strAlias & " volume to " & CStr(intVolume), "", 0, 0)

    End Sub

    Public Sub StopAndCloseSound()

        'Stop file
        intReturn = mciSendString("stop " & strAlias, "", 0, 0)

        'Close file
        intReturn = mciSendString("close " & strAlias, "", 0, 0)

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

        'Prevent cross-threading
        If _frmToPass.InvokeRequired Then
            'Invoke
            _frmToPass.Invoke(New PlaySoundDelegate(AddressOf PlaySound))
        Else
            'Open file
            intReturn = mciSendString("open " & ControlChars.Quote & _strDirectory & ControlChars.Quote & " alias " & strAlias, "", 0, 0)
            'Play file
            intReturn = mciSendString("play " & strAlias, "", 0, 0)
        End If

    End Sub

    Private Sub ClosingFile(sender As Object, e As System.Timers.ElapsedEventArgs)

        'Notes: Must have a form because MCI uses a thread and needs an invoke, otherwise no sound

        'Prevent cross-threading
        Try
            If _frmToPass.InvokeRequired Then
                'Invoke
                _frmToPass.Invoke(New ClosingSoundDelegate(AddressOf ClosingFile), New Object() {sender, e})
            Else
                'Stop timer
                sttTimer.Enabled = False
                'Stop file
                intReturn = mciSendString("stop " & strAlias, "", 0, 0)
                'Close file
                intReturn = mciSendString("close " & strAlias, "", 0, 0)
                'Remove handler
                RemoveHandler sttTimer.Elapsed, AddressOf ClosingFile
                'Empty
                sttTimer.Dispose()
                'Get thread to abort
                sttTimer = Nothing
            End If
        Catch ex As Exception
            'No debug, if errored here, most likely closed the application before this could close
        End Try

    End Sub

End Class
