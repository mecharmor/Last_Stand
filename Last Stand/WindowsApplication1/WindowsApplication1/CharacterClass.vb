'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class CharacterClass

    'Declare
    Private intFrame As Integer = 1
    Private intMovementPoint As Integer = 0

    'Bitmaps
    Public btmCharacter As Bitmap
    Public pntCharacter As Point

    'Timer
    Private sttTimer As System.Timers.Timer

    Public Sub New(intSpawnPosition As Integer, Optional blnStart As Boolean = False)

        'Preset
        btmCharacter = gbtmCharacterStand1

        'Set
        intMovementPoint = intSpawnPosition
        pntCharacter = New Point(intMovementPoint, 400)

        'Set frame timer
        sttTimer = New System.Timers.Timer(3000)
        AddHandler sttTimer.Elapsed, AddressOf Animating

        'Start
        If blnStart Then
            Start()
        End If

    End Sub

    Public Sub Start()

        'Enable timer
        sttTimer.Enabled = True

    End Sub

    Public Sub StopAndDispose()

        'Stop timer
        sttTimer.Enabled = False

        'Remove handler
        RemoveHandler sttTimer.Elapsed, AddressOf Animating

        'Dispose
        sttTimer.Dispose()

    End Sub

    Private Sub Animating(sender As Object, e As System.Timers.ElapsedEventArgs)

        'Check frame
        Select Case intFrame
            Case 1
                'Set frame
                intFrame = 2
                btmCharacter = gbtmCharacterStand2
            Case 2
                'Set frame
                intFrame = 1
                btmCharacter = gbtmCharacterStand1
            Case 3 'Interval here is still 250
                'Set frame
                intFrame = 4
                btmCharacter = gbtmCharacterShoot2
            Case 4
                'Stop and change interval
                sttTimer.Enabled = False
                sttTimer.Interval = 3000
                'Set frame
                intFrame = 1
                btmCharacter = gbtmCharacterStand1
                'Start timer again
                sttTimer.Enabled = True
        End Select

        'Do events
        Application.DoEvents()

    End Sub

    Public Sub CharacterShot()

        'Stop timer and change interval
        sttTimer.Enabled = False
        sttTimer.Interval = 250

        'Change frame immediately
        intFrame = 3
        btmCharacter = gbtmCharacterShoot1

        'Do events
        Application.DoEvents()

        'Start timer again
        sttTimer.Enabled = True

    End Sub

End Class