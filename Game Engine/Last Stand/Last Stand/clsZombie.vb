'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsZombie

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 1
    Private blnSwitch As Boolean = False
    Private intSpotX As Integer = 0
    Private intSpotY As Integer = 0
    Private _intSpeed As Integer = 25
    Private intFrameDeath As Integer = 0

    'Bitmaps
    Private btmZombie As Bitmap
    Private pntZombie As Point

    'Thread
    Private thrAnimating As System.Threading.Thread
    Private _intAnimatingDelay As Integer = 0

    'Dying
    Private blnIsDying As Boolean = False
    Private blnFirstTimeDyingPass As Boolean = False

    'Pinning
    Private blnIsPinning As Boolean = False
    Private blnFirstTimePinningPass As Boolean = False

    'Dead for painting on background
    Private blnDead As Boolean = False

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, intSpeed As Integer, Optional blnStart As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Preset
        btmZombie = gbtmZombieWalk(0)

        'Set
        _intSpeed = intSpeed

        'Set
        intSpotX = intSpawnX
        intSpotY = intSpawnY
        pntZombie = New Point(intSpotX, intSpotY)

        'Start
        If blnStart Then
            Start()
        End If

    End Sub

    Public Sub Start(Optional intAnimatingDelay As Integer = 175)

        'Set
        _intAnimatingDelay = intAnimatingDelay

        'Start thread, timed to repeat incase of too much aborting
        While Not blnStart()
        End While

    End Sub

    Private Function blnStart() As Boolean

        'Declare
        Dim blnPassed As Boolean = False

        'Start thread
        Try
            thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
            thrAnimating.Start()
            'Set
            blnPassed = True
        Catch ex As Exception
            'No debug, if got here, typed way too fast
            blnPassed = False
        End Try

        'Return
        Return blnPassed

    End Function

    Public ReadOnly Property ZombieImage() As Bitmap

        'Return
        Get
            Return btmZombie
        End Get

    End Property

    Public ReadOnly Property ZombiePoint() As Point

        'Return
        Get
            Return pntZombie
        End Get

    End Property

    Public Sub StopAndDispose()

        'Abort animating
        If thrAnimating IsNot Nothing Then
            If thrAnimating.IsAlive Then
                thrAnimating.Abort()
                thrAnimating = Nothing
            End If
        End If

    End Sub

    Private Sub Animating()

        'Continue
        While True

            'Check for first time pass dying
            If blnFirstTimeDyingPass Then
                'Set
                blnFirstTimeDyingPass = False
                'Declare
                Dim rndNumber As New Random
                'Play sound of death
                Dim udcDeathSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "/Sounds/ZombieDeath" & CStr(rndNumber.Next(1, 4)) &
                                                  ".mp3", 2000, gintSoundVolume)
                'Set frame death
                intFrameDeath = rndNumber.Next(1, 3)
                'Change frame immediately
                intFrame = 1
                Select Case intFrameDeath
                    Case 1
                        btmZombie = gbtmZombieDeath1(0)
                    Case 2
                        btmZombie = gbtmZombieDeath2(0)
                End Select
            End If

            'Check for first time pass pinning
            If blnFirstTimePinningPass Then
                'Set
                blnFirstTimePinningPass = False
                'Change frame immediately
                intFrame = 1
                btmZombie = gbtmZombiePin(0)
            End If

            'Sleep
            System.Threading.Thread.Sleep(_intAnimatingDelay)

            'Check frame, if walking, dying or pinning
            If blnIsDying Then

                'Check which frame death, sleep is 110
                Select Case intFrameDeath
                    Case 1
                        If FrameDeath(gbtmZombieDeath1) Then
                            'Exit
                            Exit While
                        End If
                    Case 2
                        If FrameDeath(gbtmZombieDeath2) Then
                            'Exit
                            Exit While
                        End If
                End Select

            Else

                'Pinning, sleep is 200
                If blnIsPinning Then

                    'Check frame
                    Select Case intFrame
                        Case 1
                            intFrame = 2
                            btmZombie = gbtmZombiePin(0)
                        Case 2
                            intFrame = 1
                            btmZombie = gbtmZombiePin(1)
                    End Select

                Else

                    'Walking, change point
                    pntZombie.X = intSpotX
                    intSpotX -= _intSpeed 'Speed they come at

                    'Check if going forward or backwards with the frames (1, 2, 3, 4) or (3, 2, 1) but becomes (1, 2, 3, 4, 3, 2, 1, 2, 3, 4, etc.)
                    If Not blnSwitch Then
                        'Check frame
                        Select Case intFrame 'Sleep here is 175 default
                            Case 1
                                intFrame = 2
                                btmZombie = gbtmZombieWalk(0)
                            Case 2
                                intFrame = 3
                                btmZombie = gbtmZombieWalk(1)
                            Case 3
                                intFrame = 4
                                btmZombie = gbtmZombieWalk(2)
                            Case 4
                                intFrame = 1
                                btmZombie = gbtmZombieWalk(3)
                                'Switch up time
                                blnSwitch = True
                        End Select
                    Else
                        'Check frame
                        Select Case intFrame
                            Case 1
                                intFrame = 2
                                btmZombie = gbtmZombieWalk(2)
                            Case 2
                                intFrame = 1
                                btmZombie = gbtmZombieWalk(1)
                                'Switch up time
                                blnSwitch = False
                        End Select
                    End If

                End If

            End If

        End While

    End Sub

    Private Function FrameDeath(btmZombieDeath() As Bitmap) As Boolean

        'Check frame
        Select Case intFrame
            Case 1
                intFrame = 2
                btmZombie = btmZombieDeath(1)
            Case 2
                intFrame = 3
                btmZombie = btmZombieDeath(2)
            Case 3
                intFrame = 4
                btmZombie = btmZombieDeath(3)
            Case 4
                intFrame = 5
                btmZombie = btmZombieDeath(4)
            Case 5
                btmZombie = btmZombieDeath(5)
                'Paint on top of the background
                blnDead = True
                'Return for exiting
                Return True
        End Select

        'Return
        Return False

    End Function

    Public ReadOnly Property IsDead() As Boolean

        'Return if ready to paint after death
        Get
            Return blnDead
        End Get

    End Property

    Public ReadOnly Property IsPinning() As Boolean

        'Return pinning or not
        Get
            Return blnIsPinning
        End Get

    End Property

    Public ReadOnly Property IsDying() As Boolean

        'Return
        Get
            Return blnIsDying
        End Get

    End Property

    Public Sub Dying()

        'Check for no instance
        If thrAnimating IsNot Nothing Then
            'Abort thread
            thrAnimating.Abort()
            'Set
            blnIsDying = True
            'Set
            blnFirstTimeDyingPass = True
            'Restart thread
            Start(110)
        End If

    End Sub

    Public Sub Pin()

        'Check for no instance
        If thrAnimating IsNot Nothing Then
            'Abort thread
            thrAnimating.Abort()
            'Set pinning
            blnIsPinning = True
            'Set
            blnFirstTimePinningPass = True
            'Restart thread
            Start(200)
        End If

    End Sub

End Class