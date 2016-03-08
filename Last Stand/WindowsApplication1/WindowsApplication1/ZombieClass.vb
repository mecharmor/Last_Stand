'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class ZombieClass

    'Declare
    Private intFrame As Integer = 1
    Private thrAnimation As System.Threading.Thread
    Private blnSwitch As Boolean = False
    Private intMovementPoint As Integer = 0
    Private _intSpeed As Integer = 25

    'Bitmaps
    Public btmZombie As Bitmap
    Public pntZombie As Point

    'Timer
    Private sttTimer As System.Timers.Timer

    'Sound
    Private udcDeathSound As SoundClass

    'Death
    Private blnAlive As Boolean = True

    Public Sub New(intSpawnPosition As Integer, Optional blnStart As Boolean = False, Optional intSpeed As Integer = 25)

        'Preset
        btmZombie = gbtmZombieWalk1

        'Set
        _intSpeed = intSpeed

        'Set
        intMovementPoint = intSpawnPosition
        pntZombie = New Point(intMovementPoint, 400)

        'Set frame timer
        sttTimer = New System.Timers.Timer(300)
        AddHandler sttTimer.Elapsed, AddressOf Animating

        'Set sound
        udcDeathSound = New SoundClass("ZombieDeath1", AppDomain.CurrentDomain.BaseDirectory & "/Sounds/ZombieDeath1.mp3")

        'Start
        If blnStart Then
            Start()
        End If

    End Sub

    Public Sub Start()

        'Enable timer
        sttTimer.Enabled = True

    End Sub

    Private Sub Animating(sender As Object, e As System.Timers.ElapsedEventArgs)

        'Check if alive before checking frame
        If blnAlive Then 'Interval is 300 in this area
            'Change point
            pntZombie = New Point(intMovementPoint, 400)
            intMovementPoint -= _intSpeed 'Speed they come at
            'Check if going forward or backwards with the frames (1, 2, 3, 4) or (3, 2, 1) but becomes (1, 2, 3, 4, 3, 2, 1, 2, 3, 4, etc.)
            If Not blnSwitch Then
                'Check frame
                Select Case intFrame
                    Case 1
                        btmZombie = gbtmZombieWalk1
                        intFrame = 2
                    Case 2
                        btmZombie = gbtmZombieWalk2
                        intFrame = 3
                    Case 3
                        btmZombie = gbtmZombieWalk3
                        intFrame = 4
                    Case 4
                        btmZombie = gbtmZombieWalk4
                        intFrame = 1
                        'Switch up time
                        blnSwitch = True
                End Select
            Else
                'Check frame
                Select Case intFrame
                    Case 1
                        btmZombie = gbtmZombieWalk3
                        intFrame = 2
                    Case 2
                        btmZombie = gbtmZombieWalk2
                        intFrame = 1
                        'Switch up time
                        blnSwitch = False
                End Select
            End If
        Else 'Interval is 75 in this area
            'Check frame
            Select Case intFrame
                Case 1
                    btmZombie = gbtmZombieDeath2
                    intFrame = 2
                Case 2
                    btmZombie = gbtmZombieDeath3
                    intFrame = 3
                Case 3
                    btmZombie = gbtmZombieDeath4
                    intFrame = 4
                Case 4
                    btmZombie = gbtmZombieDeath5
                    intFrame = 5
                Case 5
                    btmZombie = gbtmZombieDeath6
                    'Stop timer and handler
                    StopAndDispose()
            End Select

        End If

        'Do events
        Application.DoEvents()

    End Sub

    Public ReadOnly Property IsAlive() As Boolean

        'Return alive or not
        Get
            Return blnAlive
        End Get

    End Property

    Public Sub Dead()

        'Stop timer and change interval
        sttTimer.Enabled = False
        sttTimer.Interval = 75

        'Set dead
        blnAlive = False

        'Change frame immediately
        intFrame = 1
        btmZombie = gbtmZombieDeath1

        'Do events
        Application.DoEvents()

        'Play sound of death
        udcDeathSound.PlaySound(False)

        'Start timer again
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

End Class