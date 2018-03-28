'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsCharacter

    'Constants
    Private Const dblCHARACTER_STAND_DELAY As Double = 3000
    Private Const dblCHARACTER_SHOOT_DELAY As Double = 90
    Private Const dblCHARACTER_RELOAD_DELAY As Double = 85
    Private Const dblCHARACTER_RUN_DELAY As Double = 60
    Private Const intCHARACTER_RESET_POINT As Integer = 30

    'Status mode
    Public Enum eintStatusMode As Integer
        Stand = 0 'Neutral, this becomes defaulted
        Shoot = 1
        Reload = 2
        Run = 3
    End Enum

    'Declare enumeration variables
    Private intStatusModeStartToDo As eintStatusMode 'Used more in the form game
    Private intStatusModeProcessing As eintStatusMode 'Currently doing
    Private intStatusModeAboutToDo As eintStatusMode 'About to do after the processing is finished

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 0
    Private intSpotX As Integer = 0
    Private intSpotY As Integer = 0
    Private _strThisObjectName As String = ""
    Private _intBullets As Integer = 0 'Gun bullets
    Private _blnImitation As Boolean = False 'For multiplayer, ghost like properties

    'Change frame after
    Private blnChangeFrame As Boolean = False
    Private intChangeFrame As Integer = 0
    Private dblChangeInterval As Double = 0

    'Sounds
    Private _udcGunShotSound As clsSound
    Private _udcReloadingSound As clsSound
    Private _udcStepSound As clsSound
    Private _udcWaterFootStepLeftSound As clsSound
    Private _udcWaterFootStepRightSound As clsSound
    Private _udcGravelFootStepLeftSound As clsSound
    Private _udcGravelFootStepRightSound As clsSound

    'Bitmaps
    Private btmImage As Bitmap
    Private pntPoint As Point

    'Stop character from running
    Private blnStopCharacterFromRunning As Boolean = False

    'Count to remove seconds for WPM
    Private intReloadTimes As Integer = 0
    Private intRunTimes As Integer = 0

    'Send data
    Private blnSendData As Boolean = False

    'Timer
    Private tmrAnimation As New System.Timers.Timer

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, strThisObjectName As String,
                   udcReloadingSound As clsSound, udcGunShotSound As clsSound, udcStepSound As clsSound, udcWaterFootStepLeftSound As clsSound,
                   udcWaterFootStepRightSound As clsSound, udcGravelFootStepLeftSound As clsSound, udcGravelFootStepRightSound As clsSound,
                   Optional blnImitation As Boolean = False, Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set
        _strThisObjectName = strThisObjectName

        'Set
        _blnImitation = blnImitation

        'Set sounds
        _udcGunShotSound = udcGunShotSound
        _udcReloadingSound = udcReloadingSound
        _udcStepSound = udcStepSound
        _udcWaterFootStepLeftSound = udcWaterFootStepLeftSound
        _udcWaterFootStepRightSound = udcWaterFootStepRightSound
        _udcGravelFootStepLeftSound = udcGravelFootStepLeftSound
        _udcGravelFootStepRightSound = udcGravelFootStepRightSound

        'Set enumerations
        intStatusModeStartToDo = eintStatusMode.Stand
        intStatusModeAboutToDo = eintStatusMode.Stand

        'Set frame, status mode processing, and picture
        SetFrameStatusModeProcessingAndPicture(0, eintStatusMode.Stand, gabtmCharacterStandMemories(0), gabtmCharacterStandRedMemories(0),
                                               gabtmCharacterStandBlueMemories(0))

        'Set
        intSpotX = intSpawnX
        intSpotY = intSpawnY
        pntPoint = New Point(intSpotX, intSpotY)

        'Set timer
        tmrAnimation.AutoReset = True

        'Add handlers
        AddHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

        'Start
        If blnStartAnimation Then
            Start()
        End If

    End Sub

    Public Sub Start(Optional dblAnimatingDelay As Double = dblCHARACTER_STAND_DELAY)

        'Set timer delay
        tmrAnimation.Interval = dblAnimatingDelay

        'Start the animating thread
        tmrAnimation.Enabled = True

    End Sub

    Private Sub ElapsedAnimation(sender As Object, e As EventArgs)

        'Check frame
        Select Case intFrame

            Case 0 'Standing, delay here is CHARACTER_STAND_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(1, eintStatusMode.Stand, gabtmCharacterStandMemories(1), gabtmCharacterStandRedMemories(1),
                                                       gabtmCharacterStandBlueMemories(1))

            Case 1 'Standing, delay here is CHARACTER_STAND_DELAY
                'Stand
                Stand()

            Case 2 'Shooting, delay here is CHARACTER_SHOOT_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(1, eintStatusMode.Shoot, gabtmCharacterShootMemories(1), gabtmCharacterShootRedMemories(1),
                                                       gabtmCharacterShootBlueMemories(1))

            Case 3 To 22 'Reloading, delay here is CHARACTER_RELOAD_DELAY
                'Set frame
                intFrame += 1
                'Set status mode processing
                intStatusModeProcessing = eintStatusMode.Reload
                'Set picture
                Select Case _strThisObjectName
                    Case "udcCharacter"
                        btmImage = gabtmCharacterReloadMemories(intFrame - 2)
                    Case "udcCharacterOne"
                        btmImage = gabtmCharacterReloadRedMemories(intFrame - 2)
                    Case "udcCharacterTwo"
                        btmImage = gabtmCharacterReloadBlueMemories(intFrame - 2)
                End Select

            Case 23 'Reloading, delay here is CHARACTER_RELOAD_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(1, eintStatusMode.Reload, gabtmCharacterReloadMemories(21), gabtmCharacterReloadRedMemories(21),
                                                       gabtmCharacterReloadBlueMemories(21))
                'Reset bullets
                _intBullets = 0

            Case 24 To 39 'Running, sleep here is CHARACTER_RUN_DELAY
                'Set frame
                intFrame += 1
                'Set status mode processing
                intStatusModeProcessing = eintStatusMode.Run
                'Set picture
                Select Case _strThisObjectName
                    Case "udcCharacter"
                        btmImage = gabtmCharacterRunningMemories(intFrame - 24) '24 + 1 frame = 25, 25 - 24 = 1, next picture
                    Case "udcCharacterOne"
                        btmImage = Nothing
                    Case "udcCharacterTwo"
                        btmImage = Nothing
                End Select
                'Move point
                pntPoint.X = -intCHARACTER_RESET_POINT
                'Move game background
                gpntGameBackground.X -= gintCHARACTER_MOVEMENT_SPEED
                'Move zombies
                gMoveZombiesWhileRunning(gaudcZombies)
                'Play foot steps sound
                Select Case intFrame
                    Case 31, 38, 40
                        Select Case gintLevel
                            Case 1, 3
                                _udcStepSound.PlaySound(gintSoundVolume)
                            Case 2, 4
                                If intFrame = 31 Or intFrame = 40 Then
                                    _udcWaterFootStepLeftSound.PlaySound(gintSoundVolume)
                                Else
                                    _udcWaterFootStepRightSound.PlaySound(gintSoundVolume)
                                End If
                            Case 5
                                If intFrame = 31 Or intFrame = 40 Then
                                    _udcGravelFootStepLeftSound.PlaySound(gintSoundVolume)
                                Else
                                    _udcGravelFootStepRightSound.PlaySound(gintSoundVolume)
                                End If
                        End Select
                End Select
                'Check if stop running
                If intFrame = 40 Then
                    intFrame = 1
                End If

        End Select

    End Sub

    Private Sub SetFrameStatusModeProcessingAndPicture(intFrameToBe As Integer, intStatusModeProcessingToBe As eintStatusMode, btmCharacterPicture As Bitmap,
                                                       Optional btmCharacterPictureRed As Bitmap = Nothing, Optional btmCharacterPictureBlue As Bitmap = Nothing)

        'Set frame
        intFrame = intFrameToBe

        'Set processing at the moment of using a picture
        intStatusModeProcessing = intStatusModeProcessingToBe

        'Set picture
        Select Case _strThisObjectName
            Case "udcCharacter"
                btmImage = btmCharacterPicture
            Case "udcCharacterOne"
                btmImage = btmCharacterPictureRed
            Case "udcCharacterTwo"
                btmImage = btmCharacterPictureBlue
        End Select

    End Sub

    Public ReadOnly Property Image() As Bitmap

        'Return
        Get
            Return btmImage
        End Get

    End Property

    Public ReadOnly Property Point() As Point

        'Return
        Get
            Return pntPoint
        End Get

    End Property

    Public Sub StopAndDispose()

        'Disable timers
        tmrAnimation.Enabled = False

        'Stop and dispose timers
        tmrAnimation.Stop()
        tmrAnimation.Dispose()

        'Remove handlers
        RemoveHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

    End Sub

    Public Sub Shoot()

        'Set
        tmrAnimation.Interval = dblCHARACTER_SHOOT_DELAY

        'Check if not imitation
        If Not _blnImitation Then
            'Increase bullet
            _intBullets += 1
            'Check if wasted all ammo
            If _intBullets = 30 Then
                'Send data if
                If _strThisObjectName <> "udcCharacter" Then
                    PrepareSendData = True
                End If
            End If
        End If

        'Set frame, status mode processing, and picture
        SetFrameStatusModeProcessingAndPicture(2, eintStatusMode.Shoot, gabtmCharacterShootMemories(0), gabtmCharacterShootRedMemories(0),
                                               gabtmCharacterShootBlueMemories(0))

        'Default
        pntPoint.X = intSpotX

        'Set to default
        intStatusModeStartToDo = eintStatusMode.Stand

        'Play shooting sound
        _udcGunShotSound.PlaySound(gintSoundVolume) 'Sound must be after thread is started, order of operations creates smooth non glitch gameplay

    End Sub

    Public Sub Reload()

        'Stop the stop watch
        StopTypingTime() 'Only for real players

        'Set
        tmrAnimation.Interval = dblCHARACTER_RELOAD_DELAY

        'Send data
        If blnSendData Then
            'Send
            gSendData(6)
            'Set
            blnSendData = False
        End If

        'Increase
        intReloadTimes += 1

        'Set frame, status mode processing, and picture
        SetFrameStatusModeProcessingAndPicture(3, eintStatusMode.Reload, gabtmCharacterReloadMemories(0), gabtmCharacterReloadRedMemories(0),
                                               gabtmCharacterReloadBlueMemories(0))

        'Default
        pntPoint.X = intSpotX

        'Set to default
        intStatusModeStartToDo = eintStatusMode.Stand

        'Play reloading sound
        _udcReloadingSound.PlaySound(gintSoundVolume) 'Sound must be after thread is started, order of operations creates smooth non glitch gameplay

    End Sub

    Private Sub StopTypingTime()

        'Stop the stop watch
        If Not _blnImitation Then 'Only stop if real player
            gswhTimeTyped.Stop()
        End If

    End Sub

    Public Sub Stand()

        'Change delay
        tmrAnimation.Interval = dblCHARACTER_STAND_DELAY 'Forced incase it wasn't changed

        'Set frame, status mode processing, and picture
        SetFrameStatusModeProcessingAndPicture(0, eintStatusMode.Stand, gabtmCharacterStandMemories(0), gabtmCharacterStandRedMemories(0),
                                               gabtmCharacterStandBlueMemories(0))

        'Set to default
        intStatusModeStartToDo = eintStatusMode.Stand

        'Default
        pntPoint.X = intSpotX 'Forced incase looped back here from shooting

    End Sub

    Public Sub Run()

        'Stop the stop watch
        StopTypingTime() 'Only for real players

        'Set
        tmrAnimation.Interval = dblCHARACTER_RUN_DELAY

        'Increase
        intRunTimes += 1

        'Set frame, status mode processing, and picture
        SetFrameStatusModeProcessingAndPicture(24, eintStatusMode.Run, gabtmCharacterRunningMemories(0))

        'Move point
        pntPoint.X = -intCHARACTER_RESET_POINT

        'Move game background
        gpntGameBackground.X -= gintCHARACTER_MOVEMENT_SPEED

        'Move zombies
        gMoveZombiesWhileRunning(gaudcZombies)

        'Set to default
        intStatusModeStartToDo = eintStatusMode.Stand

    End Sub

    Public Property StatusModeStartToDo() As eintStatusMode

        'Return
        Get
            Return intStatusModeStartToDo
        End Get

        'Set
        Set(value As eintStatusMode)
            intStatusModeStartToDo = value
        End Set

    End Property

    Public Property StatusModeProcessing() As eintStatusMode

        'Return
        Get
            Return intStatusModeProcessing
        End Get

        'Set
        Set(value As eintStatusMode)
            intStatusModeProcessing = value
        End Set

    End Property

    Public Property StatusModeAboutToDo() As eintStatusMode

        'Return
        Get
            Return intStatusModeAboutToDo
        End Get

        'Set
        Set(value As eintStatusMode)
            intStatusModeAboutToDo = value
        End Set

    End Property

    Public Property StopCharacterFromRunning As Boolean

        'Return
        Get
            Return blnStopCharacterFromRunning
        End Get

        'Set
        Set(value As Boolean)
            blnStopCharacterFromRunning = value
        End Set

    End Property

    Public ReadOnly Property GetPictureFrame() As Integer

        'Return
        Get
            Return intFrame
        End Get

    End Property

    Public ReadOnly Property BulletsUsed() As Integer

        'Return
        Get
            Return _intBullets
        End Get

    End Property

    Public ReadOnly Property ReloadTimes() As Integer

        'Return
        Get
            Return intReloadTimes
        End Get

    End Property

    Public ReadOnly Property RunTimes() As Integer

        'Return
        Get
            Return intRunTimes
        End Get

    End Property

    Public WriteOnly Property PrepareSendData() As Boolean

        'Set
        Set(value As Boolean)
            blnSendData = value
        End Set

    End Property

End Class