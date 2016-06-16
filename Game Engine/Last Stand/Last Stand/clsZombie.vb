'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsZombie

    'Constants
    Private Const ZOMBIE_WALKING_DELAY As Integer = 160
    Private Const ZOMBIE_DYING_DELAY As Integer = 95
    Private Const ZOMBIE_PINNING_DELAY As Integer = 185

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 1
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

    'Dying
    Private blnMarkedToDie As Boolean = False
    Private blnHasPassedMarkToDie As Boolean = False
    Private blnIsDying As Boolean = False

    'Pinning
    Private blnIsPinning As Boolean = False
    Private intPinXYValueChanged As Integer = 0

    'Dead for painting on background
    Private blnDead As Boolean = False

    'Timer
    Private tmrAnimation As New System.Timers.Timer()

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

        'Set timer
        tmrAnimation.AutoReset = True

        'Add handler for the timer animation
        AddHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

        'Start
        If blnStartAnimation Then
            Start()
        End If

    End Sub

    Private Sub ElapsedAnimation(sender As Object, e As EventArgs)

        'Disable timer
        tmrAnimation.Enabled = False

        'Start thread
        thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
        thrAnimating.Start()

    End Sub

    Private Sub SetFrameAndPicture(intFrameToBe As Integer, btmZombiePicture As Bitmap, btmZombiePictureRed As Bitmap, btmZombiePictureBlue As Bitmap)

        'Set
        intFrame = intFrameToBe

        'Preset
        Select Case _strThisObjectNameCorrespondingToCharacter
            Case "udcCharacter"
                btmZombie = btmZombiePicture
            Case "udcCharacterOne"
                btmZombie = btmZombiePictureRed
            Case "udcCharacterTwo"
                btmZombie = btmZombiePictureBlue
        End Select

    End Sub

    Public Sub Start(Optional intAnimatingDelay As Integer = ZOMBIE_WALKING_DELAY, Optional blnSpawn As Boolean = True)

        'Set timer delay
        tmrAnimation.Interval = CDbl(intAnimatingDelay)

        'Set
        _blnSpawned = blnSpawn

        'Start thread
        thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
        thrAnimating.Start()

    End Sub

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
            'Wait
            While thrAnimating.IsAlive
            End While
            'Set
            thrAnimating = Nothing
            'Disable timer
            tmrAnimation.Enabled = False
            'Stop and dispose timer
            tmrAnimation.Stop()
            tmrAnimation.Dispose()
            'Remove handler
            RemoveHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation
        End If

    End Sub

    Private Sub Animating()

        'Check frame for movement
        Select Case intFrame
            Case 1 To 6
                'Walking, change point
                pntZombie.X -= _intSpeed 'Speed they come at
        End Select

        'Check frame
        Select Case intFrame

            Case 1 'Walking, delay here is ZOMBIE_WALKING_DELAY
                'Set frame and picture
                SetFrameAndPicture(2, gbtmZombieWalk(0), gbtmZombieWalkRed(0), gbtmZombieWalkBlue(0))

            Case 2 'Walking, delay here is ZOMBIE_WALKING_DELAY
                'Set frame and picture
                SetFrameAndPicture(3, gbtmZombieWalk(1), gbtmZombieWalkRed(1), gbtmZombieWalkBlue(1))

            Case 3 'Walking, delay here is ZOMBIE_WALKING_DELAY
                'Set frame and picture
                SetFrameAndPicture(4, gbtmZombieWalk(2), gbtmZombieWalkRed(2), gbtmZombieWalkBlue(2))

            Case 4 'Walking, delay here is ZOMBIE_WALKING_DELAY
                'Set frame and picture
                SetFrameAndPicture(5, gbtmZombieWalk(3), gbtmZombieWalkRed(3), gbtmZombieWalkBlue(3))

            Case 5 'Walking, delay here is ZOMBIE_WALKING_DELAY
                'Set frame and picture
                SetFrameAndPicture(6, gbtmZombieWalk(2), gbtmZombieWalkRed(2), gbtmZombieWalkBlue(2))

            Case 6 'Walking, delay here is ZOMBIE_WALKING_DELAY
                'Set frame and picture
                SetFrameAndPicture(1, gbtmZombieWalk(1), gbtmZombieWalkRed(1), gbtmZombieWalkBlue(1))

            Case 7 'Dying, delay here is ZOMBIE_DYING_DELAY
                'Declare
                Dim rndNumber As New Random
                'Play sound of death
                Dim udcDeathSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "/Sounds/ZombieDeath" & CStr(rndNumber.Next(1, 4)) &
                                                  ".mp3", 2000, gintSoundVolume)
                'Set frame death
                intFrameDeath = rndNumber.Next(1, 3)
                'Set frame, frame death, and picture
                SetFrameFrameDeathAndPicture(8, rndNumber.Next(1, 3), 0)

            Case 8 'Dying, delay here is ZOMBIE_DYING_DELAY
                'Set frame, frame death, and picture
                SetFrameFrameDeathAndPicture(9, intFrameDeath, 1)

            Case 9 'Dying, delay here is ZOMBIE_DYING_DELAY
                'Set frame, frame death, and picture
                SetFrameFrameDeathAndPicture(10, intFrameDeath, 2)

            Case 10 'Dying, delay here is ZOMBIE_DYING_DELAY
                'Set frame, frame death, and picture
                SetFrameFrameDeathAndPicture(11, intFrameDeath, 3)

            Case 11 'Dying, delay here is ZOMBIE_DYING_DELAY
                'Set frame, frame death, and picture
                SetFrameFrameDeathAndPicture(12, intFrameDeath, 4)

            Case 12 'Dying, delay here is ZOMBIE_DYING_DELAY
                'Set picture
                SetDeathTypePicture(5)
                'Set dead
                blnDead = True

            Case 13 'Pinning, delay here is ZOMBIE_PINNING_DELAY
                'Set frame and picture
                SetFrameAndPicture(14, gbtmZombiePin(0), gbtmZombiePinRed(0), gbtmZombiePinBlue(0))

            Case 14 'Pinning, delay here is ZOMBIE_PINNING_DELAY
                'Set frame and picture
                SetFrameAndPicture(13, gbtmZombiePin(1), gbtmZombiePinRed(1), gbtmZombiePinBlue(1))

        End Select

        'Enable timer, unless dead
        If Not blnDead Then
            tmrAnimation.Enabled = True
        End If

    End Sub

    Private Sub SetFrameFrameDeathAndPicture(intFrameToBe As Integer, intFrameDeathToBe As Integer, intZombieDeathIndex As Integer)

        'Set frame
        intFrame = intFrameToBe

        'Set frame death type
        intFrameDeath = intFrameDeathToBe

        'Set picture by death type
        SetDeathTypePicture(intZombieDeathIndex)

    End Sub

    Private Sub SetDeathTypePicture(intZombieDeathIndex As Integer)

        'Set picture by death type
        Select Case intFrameDeath
            Case 1
                Select Case _strThisObjectNameCorrespondingToCharacter
                    Case "udcCharacter"
                        btmZombie = gbtmZombieDeath1(intZombieDeathIndex)
                    Case "udcCharacterOne"
                        btmZombie = gbtmZombieDeathRed1(intZombieDeathIndex)
                    Case "udcCharacterTwo"
                        btmZombie = gbtmZombieDeathBlue1(intZombieDeathIndex)
                End Select
            Case 2
                Select Case _strThisObjectNameCorrespondingToCharacter
                    Case "udcCharacter"
                        btmZombie = gbtmZombieDeath2(intZombieDeathIndex)
                    Case "udcCharacterOne"
                        btmZombie = gbtmZombieDeathRed2(intZombieDeathIndex)
                    Case "udcCharacterTwo"
                        btmZombie = gbtmZombieDeathBlue2(intZombieDeathIndex)
                End Select
        End Select

    End Sub

    Public ReadOnly Property IsDead() As Boolean

        'Return
        Get
            Return blnDead
        End Get

    End Property

    Public Property PinXYValueChanged() As Integer

        'Return
        Get
            Return intPinXYValueChanged
        End Get

        'Set
        Set(value As Integer)
            intPinXYValueChanged = value
        End Set

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

        'Set
        blnIsDying = True

        'Disable timer and set frame
        DisableTimerAndSetFrame(7)

        'Restart thread
        Start(ZOMBIE_DYING_DELAY)

    End Sub

    Private Sub DisableTimerAndSetFrame(intFrameToBe As Integer)

        'Disable timer
        tmrAnimation.Enabled = False

        'Set
        intFrame = intFrameToBe

    End Sub

    Public Sub Pin()

        'Set
        blnIsPinning = True

        'Disable timer and set frame
        DisableTimerAndSetFrame(13)

        'Restart thread
        Start(ZOMBIE_PINNING_DELAY)

    End Sub

End Class