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
    Private _strThisObjectNameCorrespondingToCharacter As String = ""
    Private _blnSpawned As Boolean = False

    'Bitmaps
    Private btmZombie As Bitmap
    Private pntZombie As Point

    'Thread
    Private thrAnimating As System.Threading.Thread
    Private _intAnimatingDelay As Integer = 0

    'Dying
    Private blnMarkedToDie As Boolean = False
    Private blnHasPassedMarkToDie As Boolean = False
    Private blnIsDying As Boolean = False
    Private blnFirstTimeDyingPass As Boolean = False

    'Pinning
    Private blnIsPinning As Boolean = False
    Private blnFirstTimePinningPass As Boolean = False

    'Dead for painting on background
    Private blnDead As Boolean = False

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, intSpeed As Integer, strThisObjectNameCorrespondingToCharacter As String,
                   Optional blnStartAnimation As Boolean = False, Optional blnSpawn As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set
        _strThisObjectNameCorrespondingToCharacter = strThisObjectNameCorrespondingToCharacter

        'Set
        _intSpeed = intSpeed

        'Set
        _blnSpawned = blnSpawn

        'Set animation
        Select Case _strThisObjectNameCorrespondingToCharacter
            Case "udcCharacter"
                btmZombie = gbtmZombieWalk(0)
            Case "udcCharacterOne"
                btmZombie = gbtmZombieWalkRed(0)
            Case "udcCharacterTwo"
                btmZombie = gbtmZombieWalkBlue(0)
        End Select

        'Set
        intSpotX = intSpawnX
        intSpotY = intSpawnY
        pntZombie = New Point(intSpotX, intSpotY)

        'Start
        If blnStartAnimation Then
            Start()
        End If

    End Sub

    Public Sub Start(Optional intAnimatingDelay As Integer = 175, Optional blnSpawn As Boolean = True)

        'Set
        _blnSpawned = blnSpawn

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

    Public Property Spawned() As Boolean

        'Return
        Get
            Return _blnSpawned
        End Get

        'Set
        Set(value As Boolean)
            _blnSpawned = value
        End Set

    End Property

    Public ReadOnly Property ZombieImage() As Bitmap

        'Return
        Get
            Return btmZombie
        End Get

    End Property

    Public Property ZombiePoint() As Point

        'Return
        Get
            Return pntZombie
        End Get

        'Set
        Set(value As Point)
            pntZombie = value
        End Set

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
        While Not blnDead

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
                        Select Case _strThisObjectNameCorrespondingToCharacter
                            Case "udcCharacter"
                                btmZombie = gbtmZombieDeath1(0)
                            Case "udcCharacterOne"
                                btmZombie = gbtmZombieDeath1Red(0)
                            Case "udcCharacterTwo"
                                btmZombie = gbtmZombieDeath1Blue(0)
                        End Select
                    Case 2
                        Select Case _strThisObjectNameCorrespondingToCharacter
                            Case "udcCharacter"
                                btmZombie = gbtmZombieDeath2(0)
                            Case "udcCharacterOne"
                                btmZombie = gbtmZombieDeath2Red(0)
                            Case "udcCharacterTwo"
                                btmZombie = gbtmZombieDeath2Blue(0)
                        End Select
                End Select
            End If

            'Check for first time pass pinning
            If blnFirstTimePinningPass Then
                'Set
                blnFirstTimePinningPass = False
                'Change frame immediately
                intFrame = 1
                Select Case _strThisObjectNameCorrespondingToCharacter
                    Case "udcCharacter"
                        btmZombie = gbtmZombiePin(0)
                    Case "udcCharacterOne"
                        btmZombie = gbtmZombiePinRed(0)
                    Case "udcCharacterTwo"
                        btmZombie = gbtmZombiePinBlue(0)
                End Select
            End If

            'Sleep
            System.Threading.Thread.Sleep(_intAnimatingDelay)

            'Check frame, if walking, dying or pinning
            If blnIsDying Then

                'Check which frame death, sleep is 110
                Select Case intFrameDeath
                    Case 1
                        Select Case _strThisObjectNameCorrespondingToCharacter
                            Case "udcCharacter"
                                FrameDeath(gbtmZombieDeath1)
                            Case "udcCharacterOne"
                                FrameDeath(gbtmZombieDeath1Red)
                            Case "udcCharacterTwo"
                                FrameDeath(gbtmZombieDeath1Blue)
                        End Select
                    Case 2
                        Select Case _strThisObjectNameCorrespondingToCharacter
                            Case "udcCharacter"
                                FrameDeath(gbtmZombieDeath2)
                            Case "udcCharacterOne"
                                FrameDeath(gbtmZombieDeath2Red)
                            Case "udcCharacterTwo"
                                FrameDeath(gbtmZombieDeath2Blue)
                        End Select
                End Select

                'Check for completely dead
                If blnDead Then
                    Exit While
                End If

            Else

                'Pinning, sleep is 200
                If blnIsPinning Then

                    'Check frame
                    Select Case intFrame
                        Case 1
                            intFrame = 2
                            Select Case _strThisObjectNameCorrespondingToCharacter
                                Case "udcCharacter"
                                    btmZombie = gbtmZombiePin(0)
                                Case "udcCharacterOne"
                                    btmZombie = gbtmZombiePinRed(0)
                                Case "udcCharacterTwo"
                                    btmZombie = gbtmZombiePinBlue(0)
                            End Select
                        Case 2
                            intFrame = 1
                            Select Case _strThisObjectNameCorrespondingToCharacter
                                Case "udcCharacter"
                                    btmZombie = gbtmZombiePin(1)
                                Case "udcCharacterOne"
                                    btmZombie = gbtmZombiePinRed(1)
                                Case "udcCharacterTwo"
                                    btmZombie = gbtmZombiePinBlue(1)
                            End Select
                    End Select

                Else

                    'Walking, change point
                    pntZombie.X -= _intSpeed 'Speed they come at

                    'Check if going forward or backwards with the frames (1, 2, 3, 4) or (3, 2, 1) but becomes (1, 2, 3, 4, 3, 2, 1, 2, 3, 4, etc.)
                    If Not blnSwitch Then
                        'Check frame
                        Select Case intFrame 'Sleep here is 175 default
                            Case 1
                                intFrame = 2
                                Select Case _strThisObjectNameCorrespondingToCharacter
                                    Case "udcCharacter"
                                        btmZombie = gbtmZombieWalk(0)
                                    Case "udcCharacterOne"
                                        btmZombie = gbtmZombieWalkRed(0)
                                    Case "udcCharacterTwo"
                                        btmZombie = gbtmZombieWalkBlue(0)
                                End Select
                            Case 2
                                intFrame = 3
                                Select Case _strThisObjectNameCorrespondingToCharacter
                                    Case "udcCharacter"
                                        btmZombie = gbtmZombieWalk(1)
                                    Case "udcCharacterOne"
                                        btmZombie = gbtmZombieWalkRed(1)
                                    Case "udcCharacterTwo"
                                        btmZombie = gbtmZombieWalkBlue(1)
                                End Select
                            Case 3
                                intFrame = 4
                                Select Case _strThisObjectNameCorrespondingToCharacter
                                    Case "udcCharacter"
                                        btmZombie = gbtmZombieWalk(2)
                                    Case "udcCharacterOne"
                                        btmZombie = gbtmZombieWalkRed(2)
                                    Case "udcCharacterTwo"
                                        btmZombie = gbtmZombieWalkBlue(2)
                                End Select
                            Case 4
                                intFrame = 1
                                Select Case _strThisObjectNameCorrespondingToCharacter
                                    Case "udcCharacter"
                                        btmZombie = gbtmZombieWalk(3)
                                    Case "udcCharacterOne"
                                        btmZombie = gbtmZombieWalkRed(3)
                                    Case "udcCharacterTwo"
                                        btmZombie = gbtmZombieWalkBlue(3)
                                End Select
                                'Switch up time
                                blnSwitch = True
                        End Select
                    Else
                        'Check frame
                        Select Case intFrame
                            Case 1
                                intFrame = 2
                                Select Case _strThisObjectNameCorrespondingToCharacter
                                    Case "udcCharacter"
                                        btmZombie = gbtmZombieWalk(2)
                                    Case "udcCharacterOne"
                                        btmZombie = gbtmZombieWalkRed(2)
                                    Case "udcCharacterTwo"
                                        btmZombie = gbtmZombieWalkBlue(2)
                                End Select
                            Case 2
                                intFrame = 1
                                Select Case _strThisObjectNameCorrespondingToCharacter
                                    Case "udcCharacter"
                                        btmZombie = gbtmZombieWalk(1)
                                    Case "udcCharacterOne"
                                        btmZombie = gbtmZombieWalkRed(1)
                                    Case "udcCharacterTwo"
                                        btmZombie = gbtmZombieWalkBlue(1)
                                End Select
                                'Switch up time
                                blnSwitch = False
                        End Select
                    End If

                End If

            End If

        End While

    End Sub

    Private Sub FrameDeath(btmZombieDeath() As Bitmap)

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
        End Select

    End Sub

    Public ReadOnly Property IsDead() As Boolean

        'Return
        Get
            Return blnDead
        End Get

    End Property

    Public ReadOnly Property IsPinning() As Boolean

        'Return
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

    Public Property MarkedToDie() As Boolean

        'Notes: This zombie will only die after being rendered through the render thread.

        'Return
        Get
            Return blnMarkedToDie
        End Get

        'Set
        Set(value As Boolean)
            blnMarkedToDie = value
        End Set

    End Property

    Public Property HasPassedMarkToDie() As Boolean

        'Return
        Get
            Return blnHasPassedMarkToDie
        End Get

        'Set
        Set(value As Boolean)
            blnHasPassedMarkToDie = value
        End Set

    End Property

    Public Sub Dying()

        'Notes: Dying is more necessary than pinning, therefore this is why the subs do not look exactly the same with order of operations.

        'Check for no instance
        If thrAnimating IsNot Nothing Then
            'Abort thread
            If thrAnimating.IsAlive Then
                'Abort
                thrAnimating.Abort()
                'Set
                blnIsDying = True
                'Set
                blnFirstTimeDyingPass = True
                'Restart thread
                Start(110)
            End If
        End If

    End Sub

    Public Sub Pin()

        'Notes: Dying is more necessary than pinning, therefore this is why the subs do not look exactly the same with order of operations.

        'Check for no instance
        If thrAnimating IsNot Nothing Then
            'Abort thread
            If thrAnimating.IsAlive Then
                'Abort
                thrAnimating.Abort()
                'Set pinning
                blnIsPinning = True
                'Set
                blnFirstTimePinningPass = True
                'Restart thread
                Start(200)
            End If
        End If

    End Sub

End Class