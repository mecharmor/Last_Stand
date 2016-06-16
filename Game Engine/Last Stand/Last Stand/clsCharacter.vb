'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsCharacter

    'Constants
    Private Const CHARACTER_STAND_DELAY As Integer = 3000
    Private Const CHARACTER_SHOOT_DELAY As Integer = 175
    Private Const CHARACTER_RELOAD_DELAY As Integer = 85
    Private Const CHARACTER_RUN_DELAY As Integer = 65
    Private Const GAME_BACKGROUND_MOVING_SPEED As Integer = 60

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

    'Bitmaps
    Private btmCharacter As Bitmap
    Private pntCharacter As Point

    'Thread
    Private thrAnimating As System.Threading.Thread

    'Stop character from running
    Private blnStopCharacterFromRunning As Boolean = False

    'Reload sound
    Private udcReloadingSound As clsSound

    'Reload count
    Private intReloadTimes As Integer = 0

    'Send data
    Private blnSendData As Boolean = False

    'Timer
    Private tmrAnimation As New System.Timers.Timer()

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, strThisObjectName As String, intLevel As Integer,
                   Optional blnImitation As Boolean = False, Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set
        _strThisObjectName = strThisObjectName

        'Set
        _intLevel = intLevel

        'Set
        _blnImitation = blnImitation

        'Set enumerations
        intStatusModeStartToDo = eintStatusMode.Stand
        intStatusModeAboutToDo = eintStatusMode.Stand

        'Set frame, status mode processing, and picture
        SetFrameStatusModeProcessingAndPicture(1, eintStatusMode.Stand, gbtmCharacterStand(0), gbtmCharacterStandRed(0), gbtmCharacterStandBlue(0))

        'Set
        intSpotX = intSpawnX
        intSpotY = intSpawnY
        pntCharacter = New Point(intSpotX, intSpotY)

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

        'Start thread
        thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
        thrAnimating.Start()

    End Sub

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

        'Check frame
        Select Case intFrame

            Case 1 'Standing, delay here is CHARACTER_STAND_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(2, eintStatusMode.Stand, gbtmCharacterStand(1), gbtmCharacterStandRed(1), gbtmCharacterStandBlue(1))

            Case 2 'Standing, delay here is CHARACTER_STAND_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(1, eintStatusMode.Stand, gbtmCharacterStand(0), gbtmCharacterStandRed(0), gbtmCharacterStandBlue(0))
                'Default
                pntCharacter.X = intSpotX

            Case 3 'Shooting, delay here is CHARACTER_SHOOT_DELAY
                'Play shot sound
                Dim udcGunShotSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GunShot.mp3", 1000, gintSoundVolume)
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(4, eintStatusMode.Shoot, gbtmCharacterShoot(0), gbtmCharacterShootRed(0), gbtmCharacterShootBlue(0))
                'Default
                pntCharacter.X = intSpotX

            Case 4 'Shooting, delay here is CHARACTER_SHOOT_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(5, eintStatusMode.Shoot, gbtmCharacterShoot(1), gbtmCharacterShootRed(1), gbtmCharacterShootBlue(1))

            Case 5
                'Change delay
                tmrAnimation.Interval = CHARACTER_STAND_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(1, eintStatusMode.Stand, gbtmCharacterStand(0), gbtmCharacterStandRed(0), gbtmCharacterStandBlue(0))
                'Default
                pntCharacter.X = intSpotX

            Case 6 'Reloading, delay here is 100
                'Play reloading sound
                udcReloadingSound = New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Reloading.mp3", 3000, gintSoundVolume)
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(7, eintStatusMode.Reload, gbtmCharacterReload(0), gbtmCharacterReloadRed(0), gbtmCharacterReloadBlue(0))
                'Default
                pntCharacter.X = intSpotX

            Case 7 To 26 'Reloading, delay here is 100
                'Set frame
                intFrame += 1
                'Set status mode processing
                intStatusModeProcessing = eintStatusMode.Reload
                'Set picture
                Select Case _strThisObjectName
                    Case "udcCharacter"
                        btmCharacter = gbtmCharacterReload(intFrame - 7)
                    Case "udcCharacterOne"
                        btmCharacter = gbtmCharacterReloadRed(intFrame - 7)
                    Case "udcCharacterTwo"
                        btmCharacter = gbtmCharacterReloadBlue(intFrame - 7)
                End Select

            Case 27 'Reloading, delay here is CHARACTER_RELOAD_DELAY
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(5, eintStatusMode.Reload, gbtmCharacterReload(21), gbtmCharacterReloadRed(21), gbtmCharacterReloadBlue(21))
                'Reset bullets
                _intBullets = 0
                'Reset key press event bullets
                gintBullets = 0

            Case 28
                'Set frame, status mode processing, and picture
                SetFrameStatusModeProcessingAndPicture(29, eintStatusMode.Run, gbtmCharacterRunning(0))
                'Move point
                pntCharacter.X = -GAME_BACKGROUND_MOVING_SPEED
                'Move game background
                gpntGameBackground.X -= gMOVEMENTSPEED
                'Move zombies
                gMoveZombiesWhileRunning(gaudcZombies)

            Case 29 To 44 'Running, sleep here is CHARACTER_RUN_DELAY
                'Set frame
                intFrame += 1
                'Set status mode processing
                intStatusModeProcessing = eintStatusMode.Run
                'Set picture
                Select Case _strThisObjectName
                    Case "udcCharacter"
                        btmCharacter = gbtmCharacterRunning(intFrame - 29) '30-29 = 1 in the array
                    Case "udcCharacterOne"
                        btmCharacter = Nothing
                    Case "udcCharacterTwo"
                        btmCharacter = Nothing
                End Select
                'Move point
                pntCharacter.X = -GAME_BACKGROUND_MOVING_SPEED
                'Play foot steps sound
                Select Case intFrame
                    Case 35, 42, 44
                        Select Case _intLevel
                            Case 1, 3
                                Dim udcStepSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Step.mp3", 1000, gintSoundVolume)
                            Case 2, 4
                                If intFrame = 35 Or intFrame = 44 Then
                                    Dim udcWaterFootStepLeftSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory &
                                                                                  "Sounds\WaterFootStepLeft.mp3", 1000, gintSoundVolume)
                                Else
                                    Dim udcWaterFootStepRightSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory &
                                                                                   "Sounds\WaterFootStepRight.mp3", 1000, gintSoundVolume)
                                End If
                            Case 5
                                If intFrame = 35 Or intFrame = 44 Then
                                    Dim udcGravelFootStepLeftSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory &
                                                                                   "Sounds\GravelFootStepLeft.mp3", 2000, gintSoundVolume)
                                Else
                                    Dim udcGravelFootStepRightSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory &
                                                                                    "Sounds\GravelFootStepRight.mp3", 1000, gintSoundVolume)
                                End If
                        End Select
                End Select
                'Move game background
                gpntGameBackground.X -= gMOVEMENTSPEED
                'Move zombies
                gMoveZombiesWhileRunning(gaudcZombies)
                'Check if stop running
                If intFrame = 45 Then
                    intFrame = 5
                End If

        End Select

        'Enable timer
        tmrAnimation.Enabled = True

    End Sub

    Public Sub Shoot()

        'Disable timer, set frame, and set status modes
        DisableTimerSetFrameAndStatusModes(3, eintStatusMode.Shoot)

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

        'Restart thread
        Start(CHARACTER_SHOOT_DELAY)

    End Sub

    Private Sub DisableTimerSetFrameAndStatusModes(intFrameToBe As Integer, intStatusModeProcessingToBe As eintStatusMode)

        'Disable timer
        tmrAnimation.Enabled = False

        'Set to default
        intStatusModeStartToDo = eintStatusMode.Stand

        'Set enumeration
        intStatusModeProcessing = intStatusModeProcessingToBe

        'Set
        intFrame = intFrameToBe

    End Sub

    Public Sub Reload()

        'Disable timer, set frame, and set status modes
        DisableTimerSetFrameAndStatusModes(6, eintStatusMode.Reload)

        'Send data
        If blnSendData Then
            'Send
            gSendData("5|")
            'Set
            blnSendData = False
        End If

        'Increase
        intReloadTimes += 1

        'Restart thread
        Start(CHARACTER_RELOAD_DELAY)

    End Sub

    Public Sub Stand()

        'Disable timer, set frame, and set status modes
        DisableTimerSetFrameAndStatusModes(2, eintStatusMode.Stand)

        'Restart thread
        Start() 'Default CHARACTER_STAND_DELAY

    End Sub

    Public Sub Run()

        'Disable timer, set frame, and set status modes
        DisableTimerSetFrameAndStatusModes(28, eintStatusMode.Run)

        'Restart thread
        Start(CHARACTER_RUN_DELAY)

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

    Public Sub StopReloadingSound()

        'Stop sound
        If udcReloadingSound IsNot Nothing Then
            udcReloadingSound.StopAndCloseSound()
            udcReloadingSound = Nothing
        End If

    End Sub

End Class