'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsCharacter

    'Constants
    Private Const CHARACTER_STAND_DELAY As Integer = 3000
    Private Const CHARACTER_SHOOT_DELAY As Integer = 175
    Private Const CHARACTER_RELOAD_DELAY As Integer = 85
    Private Const CHARACTER_RUN_DELAY As Integer = 60
    Private Const CHARACTER_RESET_POINT As Integer = 60

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
    Private _intLevel As Integer = 1 'Starting level
    Private _blnImitation As Boolean = False 'For multiplayer, ghost like properties
    Private _blnSpawned As Boolean = False

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

    'Ending thread
    Private blnThreadDisposing As Boolean = False

    'Bitmaps
    Private btmCharacter As Bitmap
    Private pntCharacter As Point

    'Thread
    Private thrAnimating As System.Threading.Thread

    'Stop character from running
    Private blnStopCharacterFromRunning As Boolean = False

    'Reload count
    Private intReloadTimes As Integer = 0

    'Send data
    Private blnSendData As Boolean = False

    'Timer
    Private tmrAnimation As New System.Timers.Timer

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, strThisObjectName As String, intLevel As Integer,
                   udcReloadingSound As clsSound, udcGunShotSound As clsSound, udcStepSound As clsSound, udcWaterFootStepLeftSound As clsSound,
                   udcWaterFootStepRightSound As clsSound, udcGravelFootStepLeftSound As clsSound, udcGravelFootStepRightSound As clsSound,
                   Optional blnImitation As Boolean = False, Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set
        _strThisObjectName = strThisObjectName

        'Set
        _intLevel = intLevel

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
        SetFrameStatusModeProcessingAndPicture(1, eintStatusMode.Stand, gabtmCharacterStandMemories(0), gabtmCharacterStandRedMemories(0),
                                               gabtmCharacterStandBlueMemories(0))

        'Set
        intSpotX = intSpawnX
        intSpotY = intSpawnY
        pntCharacter = New Point(intSpotX, intSpotY)

        'Set timer
        tmrAnimation.AutoReset = True

        'Add handlers
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
        If Not blnThreadDisposing Then
            thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
            thrAnimating.Start()
        End If

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
                btmCharacter = btmCharacterPicture
            Case "udcCharacterOne"
                btmCharacter = btmCharacterPictureRed
            Case "udcCharacterTwo"
                btmCharacter = btmCharacterPictureBlue
        End Select

    End Sub

    Public Sub Start(Optional intAnimatingDelay As Integer = CHARACTER_STAND_DELAY)

        'Set timer delay
        tmrAnimation.Interval = CDbl(intAnimatingDelay)

        'Start the animating thread
        thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
        thrAnimating.Start()

        'Set
        _blnSpawned = True 'This must be after the thread is created for timing pro

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

    Public ReadOnly Property CharacterImage() As Bitmap

        'Return
        Get
            Return btmCharacter
        End Get

    End Property

    Public ReadOnly Property CharacterPoint() As Point

        'Return
        Get
            Return pntCharacter
        End Get

    End Property

    Public Sub StopAndDispose()

        'Set
        blnThreadDisposing = True

        'Disable timers
        tmrAnimation.Enabled = False

        'Stop and dispose timers
        tmrAnimation.Stop()
        tmrAnimation.Dispose()

        'Remove handlers
        RemoveHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

        'Abort animating
        If thrAnimating IsNot Nothing Then
            'Wait
            While thrAnimating.IsAlive
            End While
            'Set
            thrAnimating = Nothing
        End If

    End Sub

    Private Sub Animating()

        'Check frame
        Select Case intFrame

            Case 1 'Standing, delay here is CHARACTER_STAND_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(2, eintStatusMode.Stand, gabtmCharacterStandMemories(1), gabtmCharacterStandRedMemories(1),
                                                       gabtmCharacterStandBlueMemories(1))

            Case 2 'Standing, delay here is CHARACTER_STAND_DELAY
                'Stand
                StandSub() 'This must be here, because the frame could be cycled to view standing

            Case 3 'Shooting, delay here is CHARACTER_SHOOT_DELAY
                'Not used here

            Case 4 'Shooting, delay here is CHARACTER_SHOOT_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(5, eintStatusMode.Shoot, gabtmCharacterShootMemories(1), gabtmCharacterShootRedMemories(1),
                                                       gabtmCharacterShootBlueMemories(1))

            Case 5
                'Change delay
                tmrAnimation.Interval = CHARACTER_STAND_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(1, eintStatusMode.Stand, gabtmCharacterStandMemories(0), gabtmCharacterStandRedMemories(0),
                                                       gabtmCharacterStandBlueMemories(0))
                'Default
                pntCharacter.X = intSpotX

            Case 6 'Reloading, delay here is CHARACTER_RELOAD_DELAY
                'Not used here

            Case 7 To 26 'Reloading, delay here is CHARACTER_RELOAD_DELAY
                'Set frame
                intFrame += 1
                'Set status mode processing
                intStatusModeProcessing = eintStatusMode.Reload
                'Set picture
                Select Case _strThisObjectName
                    Case "udcCharacter"
                        btmCharacter = gabtmCharacterReloadMemories(intFrame - 7)
                    Case "udcCharacterOne"
                        btmCharacter = gabtmCharacterReloadRedMemories(intFrame - 7)
                    Case "udcCharacterTwo"
                        btmCharacter = gabtmCharacterReloadBlueMemories(intFrame - 7)
                End Select

            Case 27 'Reloading, delay here is CHARACTER_RELOAD_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(5, eintStatusMode.Reload, gabtmCharacterReloadMemories(21), gabtmCharacterReloadRedMemories(21),
                                                       gabtmCharacterReloadBlueMemories(21))
                'Reset bullets
                _intBullets = 0
                'Reset key press event bullets
                gintBullets = 0

            Case 28 'Run, delay here is CHARACTER_RUN_DELAY
                'Not used here

            Case 29 To 44 'Running, sleep here is CHARACTER_RUN_DELAY
                'Set frame
                intFrame += 1
                'Set status mode processing
                intStatusModeProcessing = eintStatusMode.Run
                'Set picture
                Select Case _strThisObjectName
                    Case "udcCharacter"
                        btmCharacter = gabtmCharacterRunningMemories(intFrame - 29) '30-29 = 1 in the array
                    Case "udcCharacterOne"
                        btmCharacter = Nothing
                    Case "udcCharacterTwo"
                        btmCharacter = Nothing
                End Select
                'Move point
                pntCharacter.X = -CHARACTER_RESET_POINT
                'Move game background
                gpntGameBackground.X -= gCHARACTER_MOVEMENT_SPEED
                'Move zombies
                gMoveZombiesWhileRunning(gaudcZombies)
                'Play foot steps sound
                Select Case intFrame
                    Case 35, 42, 44
                        Select Case _intLevel
                            Case 1, 3
                                _udcStepSound.PlaySound(gintSoundVolume)
                            Case 2, 4
                                If intFrame = 35 Or intFrame = 44 Then
                                    _udcWaterFootStepLeftSound.PlaySound(gintSoundVolume)
                                Else
                                    _udcWaterFootStepRightSound.PlaySound(gintSoundVolume)
                                End If
                            Case 5
                                If intFrame = 35 Or intFrame = 44 Then
                                    _udcGravelFootStepLeftSound.PlaySound(gintSoundVolume)
                                Else
                                    _udcGravelFootStepRightSound.PlaySound(gintSoundVolume)
                                End If
                        End Select
                End Select
                'Check if stop running
                If intFrame = 45 Then
                    intFrame = 5
                End If

        End Select

        'Enable timer, unless need to avoid or dispose
        If Not blnThreadDisposing Then
            tmrAnimation.Enabled = True
        End If

    End Sub

    Private Sub StandSub()

        'Set frame, status mode processing, and picture
        SetFrameStatusModeProcessingAndPicture(1, eintStatusMode.Stand, gabtmCharacterStandMemories(0), gabtmCharacterStandRedMemories(0),
                                               gabtmCharacterStandBlueMemories(0))

        'Default
        pntCharacter.X = intSpotX

    End Sub

    Public Sub Shoot()

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
        SetFrameStatusModeProcessingAndPicture(4, eintStatusMode.Shoot, gabtmCharacterShootMemories(0), gabtmCharacterShootRedMemories(0),
                                               gabtmCharacterShootBlueMemories(0))

        'Default
        pntCharacter.X = intSpotX

        'Change status and interval
        ChangeStatusAndInterval(CHARACTER_SHOOT_DELAY)

        'Play shooting sound
        _udcGunShotSound.PlaySound(gintSoundVolume) 'Sound must be after thread is started, order of operations creates smooth non glitch gameplay

    End Sub

    Private Sub ChangeStatusAndInterval(intFrameDelay As Integer)

        'Set to default
        intStatusModeStartToDo = eintStatusMode.Stand

        'Set
        tmrAnimation.Enabled = False

        'Set
        tmrAnimation.Interval = CDbl(intFrameDelay)

        'Set
        tmrAnimation.Enabled = True

    End Sub

    Public Sub Reload()

        'Send data
        If blnSendData Then
            'Send
            gSendData(6, "")
            'Set
            blnSendData = False
        End If

        'Increase
        intReloadTimes += 1

        'Set frame, status mode processing, and picture
        SetFrameStatusModeProcessingAndPicture(7, eintStatusMode.Reload, gabtmCharacterReloadMemories(0), gabtmCharacterReloadRedMemories(0),
                                               gabtmCharacterReloadBlueMemories(0))

        'Default
        pntCharacter.X = intSpotX

        'Change status and interval
        ChangeStatusAndInterval(CHARACTER_RELOAD_DELAY)

        'Play reloading sound
        _udcReloadingSound.PlaySound(gintSoundVolume) 'Sound must be after thread is started, order of operations creates smooth non glitch gameplay

    End Sub

    Public Sub Stand()

        'Stand
        StandSub()

        'Change status and interval
        ChangeStatusAndInterval(CHARACTER_STAND_DELAY)

    End Sub

    Public Sub Run()

        'Set frame, status mode processing, and picture
        SetFrameStatusModeProcessingAndPicture(29, eintStatusMode.Run, gabtmCharacterRunningMemories(0))

        'Move point
        pntCharacter.X = -CHARACTER_RESET_POINT

        'Move game background
        gpntGameBackground.X -= gCHARACTER_MOVEMENT_SPEED

        'Move zombies
        gMoveZombiesWhileRunning(gaudcZombies)

        'Change status and interval
        ChangeStatusAndInterval(CHARACTER_RUN_DELAY)

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

    Public WriteOnly Property PrepareSendData() As Boolean

        'Set
        Set(value As Boolean)
            blnSendData = value
        End Set

    End Property

End Class