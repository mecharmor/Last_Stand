'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsSound

    'Declare function
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String,
                                                                                   ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    'Declare
    Private _frmToPass As Form
    Private _strAlias As String = ""
    Private intReturn As Integer = 0
    Private _intNumberOfSounds As Integer = 0

    'Notes: This class needs invokes because of delegation within threads. MCISendString seems to run in a thread and won't work without invoke when in another thread.

    Public Sub New(frmToPass As Form, strDirectory As String, intNumberOfSounds As Integer)

        'Notes: Number of sounds will be 1 for repeating sounds. Sounds that must be continuously heard and new ones played, must use more than 1 and will not repeat.

        'Set
        _frmToPass = frmToPass

        'Set
        _strAlias = strGetAlias(strDirectory)

        'Increase counter
        gintAlias += 1

        'Set
        _intNumberOfSounds = intNumberOfSounds

        'Declare
        Dim intIndex As Integer = 0 'Lambda reason

        'Prepare to load sounds
        If IO.File.Exists(strDirectory) Then
            'Loop to open all sounds
            For intLoop As Integer = 0 To _intNumberOfSounds - 1
                'Set
                intIndex = intLoop
                'Open individual
                _frmToPass.Invoke(Sub() intReturn = mciSendString("open " & ControlChars.Quote & strDirectory & ControlChars.Quote & " alias " & _strAlias &
                                  CStr(intIndex), "", 0, 0))
            Next
        Else
            'Sound is missing, close program
            gCloseApplicationWithErrorMessage("The " & strDirectory & " file is missing. This application will close now.")
        End If

    End Sub

    Public Sub StopAndCloseSound()

        'Declare
        Dim intIndex As Integer = 0 'Lambda reason

        'Stop and close all sounds
        For intLoop As Integer = 0 To _intNumberOfSounds - 1
            'Set
            intIndex = intLoop
            'Stop
            _frmToPass.Invoke(Sub() intReturn = mciSendString("stop " & _strAlias & CStr(intIndex), "", 0, 0))
            'Close
            _frmToPass.Invoke(Sub() intReturn = mciSendString("close " & _strAlias & CStr(intIndex), "", 0, 0))
        Next

    End Sub

    Private Function strGetAlias(strDirectory As String) As String

        'Notes: Looks like "X:\folder\folder\folder\filename.mp3"

        'Declare
        Dim astrTemp1() As String = Split(strDirectory, "\")
        Dim astrTemp2() As String = Split(astrTemp1(astrTemp1.GetUpperBound(0)), ".")

        'Return
        Return astrTemp2(0) & "_" & CStr(gintAlias) & "_" 'Looks like "filename_i_" as end result with mciSendString

    End Function

    Public Sub ChangeVolumeWhileSoundIsPlaying()

        'Set volume
        _frmToPass.Invoke(Sub() intReturn = mciSendString("setaudio " & _strAlias & CStr(_intNumberOfSounds - 1) & " volume to " & CStr(gintSoundVolume), "", 0, 0))

    End Sub

    Public Sub PlaySound(intVolume As Integer, Optional blnRepeat As Boolean = False)

        'Check if repeating sound or not
        If blnRepeat Then

            'Set volume
            _frmToPass.Invoke(Sub() intReturn = mciSendString("setaudio " & _strAlias & CStr(_intNumberOfSounds - 1) & " volume to " & CStr(intVolume), "", 0, 0))

            'Seek
            _frmToPass.Invoke(Sub() intReturn = mciSendString("seek " & _strAlias & CStr(_intNumberOfSounds - 1) & " to start", "", 0, 0))

            'Repeating sound
            _frmToPass.Invoke(Sub() intReturn = mciSendString("play " & _strAlias & CStr(_intNumberOfSounds - 1) & " repeat", "", 0, 0))

        Else

            'Declare
            Dim intIndex As Integer = 0 'Lambda reason

            'Loop to play sound
            For intLoop As Integer = 0 To _intNumberOfSounds - 1
                'Check if sound is not playing
                If Not SoundIsPlaying(intLoop) Then
                    'Set
                    intIndex = intLoop
                    'Set volume
                    _frmToPass.Invoke(Sub() intReturn = mciSendString("setaudio " & _strAlias & CStr(intIndex) & " volume to " & CStr(intVolume), "", 0, 0))
                    'Seek
                    _frmToPass.Invoke(Sub() intReturn = mciSendString("seek " & _strAlias & CStr(intIndex) & " to start", "", 0, 0))
                    'Play
                    _frmToPass.Invoke(Sub() intReturn = mciSendString("play " & _strAlias & CStr(intIndex), "", 0, 0))
                    'Exit
                    Exit For
                End If
            Next

        End If

    End Sub

    Private Function SoundIsPlaying(intIndex As Integer) As Boolean

        'Declare
        Dim strReturn As String = Space(128) 'Buffered string, has to be used this way

        'Check sound status
        _frmToPass.Invoke(Sub() intReturn = mciSendString("status " & _strAlias & CStr(intIndex) & " mode", strReturn, 128, 0))

        'Fix the buffered string into a regular string
        strReturn = Replace(strReturn, Chr(0), "").Trim

        'Return
        If strReturn = "stopped" Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Sub StopSound()

        'Declare
        Dim intIndex As Integer = 0 'Lambda reason

        'Stop all sounds
        For intLoop As Integer = 0 To _intNumberOfSounds - 1
            'Set
            intIndex = intLoop
            'Stop
            _frmToPass.Invoke(Sub() intReturn = mciSendString("stop " & _strAlias & CStr(intIndex), "", 0, 0))
        Next

    End Sub

End Class