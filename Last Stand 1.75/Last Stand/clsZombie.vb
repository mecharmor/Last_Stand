'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsZombie

    'Constants
    Private Const dblZOMBIE_WALKING_DELAY As Double = 160
    Private Const dblZOMBIE_DYING_DELAY As Double = 95
    Private Const dblZOMBIE_PINNING_DELAY As Double = 185

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 0
    Private intSpotX As Integer = 0
    Private intSpotY As Integer = 0
    Private _intSpeed As Integer = 12
    Private intFrameDeath As Integer = 0
    Private _strThisObjectNameCorrespondingToCharacter As String = ""
    Private _blnImitation As Boolean = False 'For multiplayer, ghost like properties
    Private _blnSpawned As Boolean = False

    'Sounds
    Private _audcZombieDeathSounds() As clsSound
    Private _udcWaterSplashSound As clsSound

    'Bitmaps
    Private btmImage As Bitmap
    Private pntPoint As Point

    'Texture index for OpenGL
    Private intTextureIndex As Integer

    'Dying
    Private blnIsDying As Boolean = False

    'Pinning
    Private blnIsPinning As Boolean = False

    'Dead for painting on background
    Private blnDead As Boolean = False

    'Timer
    Private tmrAnimation As New System.Timers.Timer

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, intSpeed As Integer,
                   strThisObjectNameCorrespondingToCharacter As String,
                   audcZombieDeathSounds() As clsSound, udcWaterSplashSound As clsSound,
                   blnAnimationStartRandom As Boolean, Optional blnImitation As Boolean = False,
                   Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set
        _strThisObjectNameCorrespondingToCharacter = strThisObjectNameCorrespondingToCharacter

        'Set
        _blnImitation = blnImitation

        'Set
        _intSpeed = intSpeed

        'Set zombie death sounds
        _audcZombieDeathSounds = audcZombieDeathSounds

        'Set splash sounds
        _udcWaterSplashSound = udcWaterSplashSound

        'Set animation
        If blnAnimationStartRandom Then
            'Get a random frame
            Dim intRandomAnimation As Integer = gGetRandomNumber(0, 5)
            Select Case intRandomAnimation
                Case 0 To 3
                    SetFrameAndPicture(intRandomAnimation + 1, gabtmZombieWalkMemories(intRandomAnimation), 0 + intRandomAnimation,
                                       gabtmZombieWalkRedMemories(intRandomAnimation), 18 + intRandomAnimation,
                                       gabtmZombieWalkBlueMemories(intRandomAnimation), 36 + intRandomAnimation)
                Case 4
                    SetFrameAndPicture(5, gabtmZombieWalkMemories(2), 2, gabtmZombieWalkRedMemories(2), 20, gabtmZombieWalkBlueMemories(2), 38)
                Case 5
                    SetFrameAndPicture(0, gabtmZombieWalkMemories(1), 1, gabtmZombieWalkRedMemories(1), 19, gabtmZombieWalkBlueMemories(1), 37)
            End Select
        Else
            'First frame
            SetFrameAndPicture(1, gabtmZombieWalkMemories(0), 0, gabtmZombieWalkRedMemories(0), 18, gabtmZombieWalkBlueMemories(0), 36)
        End If

        'Set
        intSpotX = intSpawnX
        intSpotY = intSpawnY
        pntPoint = New Point(intSpotX, intSpotY)

        'Set timer
        tmrAnimation.AutoReset = True

        'Add handler for the timer animation
        AddHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

        'Start
        If blnStartAnimation Then
            Start(dblZOMBIE_WALKING_DELAY)
        End If

    End Sub

    Public Sub Start(Optional dblAnimatingDelay As Double = dblZOMBIE_WALKING_DELAY)

        'Set timer delay
        tmrAnimation.Interval = dblAnimatingDelay

        'Start the animating thread
        tmrAnimation.Enabled = True

        'Set
        _blnSpawned = True 'This must be after the thread is created for timing problems

    End Sub

    Private Sub ElapsedAnimation(sender As Object, e As EventArgs)

        'Check frame
        Select Case intFrame

            Case 0 To 3 'Walking, delay here is ZOMBIE_WALKING_DELAY
                'Set frame and picture
                SetFrameAndPicture(intFrame + 1, gabtmZombieWalkMemories(intFrame), 0 + intFrame, gabtmZombieWalkRedMemories(intFrame),
                                   18 + intFrame, gabtmZombieWalkBlueMemories(intFrame), 36 + intFrame)
                'Make zombie walking forward
                ZombieWalkForward()

            Case 4 'Walking, delay here is ZOMBIE_WALKING_DELAY
                'Set frame and picture
                SetFrameAndPicture(5, gabtmZombieWalkMemories(2), 2, gabtmZombieWalkRedMemories(2), 20, gabtmZombieWalkBlueMemories(2), 38)
                'Make zombie walking forward
                ZombieWalkForward()

            Case 5 'Walking, delay here is ZOMBIE_WALKING_DELAY
                'Set frame and picture
                SetFrameAndPicture(0, gabtmZombieWalkMemories(1), 1, gabtmZombieWalkRedMemories(1), 19, gabtmZombieWalkBlueMemories(1), 37)
                'Make zombie walking forward
                ZombieWalkForward()

            Case 6 To 9 'Dying, delay here is ZOMBIE_DYING_DELAY
                'Set frame, frame death, and picture
                SetFrameAndDeathPicture(intFrame + 1, intFrame - 6)

            Case 10 'Dying, delay here is ZOMBIE_DYING_DELAY
                'Set picture
                SetDeathTypePicture(5)
                'Play special effect
                Select Case gintLevel
                    Case 2
                        'Play sound
                        _udcWaterSplashSound.PlaySound(gintSoundVolume)
                End Select
                'Set dead
                blnDead = True
                'Disable timer completely
                StopAndDisposeTimer()

            Case 11 'Pinning, delay here is ZOMBIE_PINNING_DELAY
                'Set frame and picture
                SetFrameAndPicture(12, gabtmZombiePinMemories(0), 16, gabtmZombiePinRedMemories(0), 34,
                                   gabtmZombiePinBlueMemories(0), 52) 'Must be here, comes back in a cycle

            Case 12 'Pinning, delay here is ZOMBIE_PINNING_DELAY
                'Set frame and picture
                SetFrameAndPicture(11, gabtmZombiePinMemories(1), 17, gabtmZombiePinRedMemories(1), 35, gabtmZombiePinBlueMemories(1), 53)

        End Select

    End Sub

    Private Sub ZombieWalkForward()

        'Walking, change point, only change if not a ghost like property
        If Not _blnImitation Then
            pntPoint.X -= _intSpeed 'Speed they come at
        End If

    End Sub

    Private Sub SetFrameAndPicture(intFrameToBe As Integer, btmZombiePicture As Bitmap, intTextureIndexToBe As Integer,
                                   btmZombiePictureRed As Bitmap, intTextureIndexToBeRed As Integer,
                                   btmZombiePictureBlue As Bitmap, intTextureIndexToBeBlue As Integer)

        'Set
        intFrame = intFrameToBe

        'Preset
        Select Case _strThisObjectNameCorrespondingToCharacter
            Case "udcCharacter"
                btmImage = btmZombiePicture
                intTextureIndex = intTextureIndexToBe
            Case "udcCharacterOne"
                btmImage = btmZombiePictureRed
                intTextureIndex = intTextureIndexToBeRed
            Case "udcCharacterTwo"
                btmImage = btmZombiePictureBlue
                intTextureIndex = intTextureIndexToBeBlue
        End Select

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

    Public ReadOnly Property Image() As Bitmap

        'Return
        Get
            Return btmImage
        End Get

    End Property

    Public Property Point() As Point

        'Return
        Get
            Return pntPoint
        End Get

        'Set
        Set(value As Point)
            pntPoint = value
        End Set

    End Property

    Public ReadOnly Property TextureIndex() As Integer

        'Return
        Get
            Return intTextureIndex
        End Get

    End Property

    Public Sub StopAndDisposeTimer()

        'Stop and dispose timer variable
        gStopAndDisposeTimerVariable(tmrAnimation)

        'Remove handler
        RemoveHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

    End Sub

    Private Sub SetFrameAndDeathPicture(intFrameToBe As Integer, intZombieDeathIndex As Integer)

        'Set frame
        intFrame = intFrameToBe

        'Set picture by death type
        SetDeathTypePicture(intZombieDeathIndex)

    End Sub

    Private Sub SetDeathTypePicture(intZombieDeathIndex As Integer)

        'Set picture by death type
        Select Case intFrameDeath
            Case 0
                Select Case _strThisObjectNameCorrespondingToCharacter
                    Case "udcCharacter"
                        btmImage = gabtmZombieDeath1Memories(intZombieDeathIndex)
                        intTextureIndex = 4 + intZombieDeathIndex
                    Case "udcCharacterOne"
                        btmImage = gabtmZombieDeathRed1Memories(intZombieDeathIndex)
                        intTextureIndex = 22 + intZombieDeathIndex
                    Case "udcCharacterTwo"
                        btmImage = gabtmZombieDeathBlue1Memories(intZombieDeathIndex)
                        intTextureIndex = 40 + intZombieDeathIndex
                End Select
            Case 1
                Select Case _strThisObjectNameCorrespondingToCharacter
                    Case "udcCharacter"
                        btmImage = gabtmZombieDeath2Memories(intZombieDeathIndex)
                        intTextureIndex = 10 + intZombieDeathIndex
                    Case "udcCharacterOne"
                        btmImage = gabtmZombieDeathRed2Memories(intZombieDeathIndex)
                        intTextureIndex = 28 + intZombieDeathIndex
                    Case "udcCharacterTwo"
                        btmImage = gabtmZombieDeathBlue2Memories(intZombieDeathIndex)
                        intTextureIndex = 46 + intZombieDeathIndex
                End Select
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

    Public Sub Dying()

        'Disable
        tmrAnimation.Enabled = False

        'Set
        tmrAnimation.Interval = dblZOMBIE_DYING_DELAY

        'Set
        blnIsDying = True

        'Set frame death
        intFrameDeath = gGetRandomNumber(0, 1)

        'Set frame, frame death, and picture
        SetFrameAndDeathPicture(6, 0)

        'Play zombie death sound
        _audcZombieDeathSounds(gGetRandomNumber(0, 1)).PlaySound(gintSoundVolume)

        'Enable
        tmrAnimation.Enabled = True

    End Sub

    Public Sub Pin()

        'Disable
        tmrAnimation.Enabled = False

        'Set
        tmrAnimation.Interval = dblZOMBIE_PINNING_DELAY

        'Set
        blnIsPinning = True

        'Set frame and picture
        SetFrameAndPicture(12, gabtmZombiePinMemories(0), 16, gabtmZombiePinRedMemories(0), 34, gabtmZombiePinBlueMemories(0), 52)

        'Enable
        tmrAnimation.Enabled = True

    End Sub

End Class