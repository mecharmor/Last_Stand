'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsCharacter

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 1
    Private intSpotX As Integer = 0
    Private intSpotY As Integer = 0
    Private _strThisObjectName As String = ""
    Private _intBullets As Integer = 0 'Gun bullets
    Private _blnImitation As Boolean = False 'For multiplayer, ghost like properties

    'For error preventing
    Private blnAborted As Boolean = False
    Private blnKeepUsingAnimatingThread As Boolean = True

    'Bitmaps
    Private btmCharacter As Bitmap
    Private pntCharacter As Point

    'Thread
    Private thrAnimating As System.Threading.Thread
    Private _intAnimatingDelay As Integer = 0

    'Shooting
    Private blnIsShooting As Boolean = False
    Private blnFirstTimeShootingPass As Boolean = False

    'Reloading
    Private blnIsReloading As Boolean = False
    Private blnNeedsToShowReloading As Boolean = False
    Private blnFirstTimeReloadingPass As Boolean = False

    'Running
    Private blnIsRunning As Boolean = False
    Private blnFirstTimeRunningPass As Boolean = False
    Private blnPrepareToRun As Boolean = False
    Private blnEndOfLevel As Boolean = False

    'Reload sound
    Private udcReloadingSound As clsSound

    'Reload count
    Private intReloadTimes As Integer = 0

    'Send data
    Private blnSendData As Boolean = False

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, strThisObjectName As String, Optional blnImitation As Boolean = False,
                   Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set
        _strThisObjectName = strThisObjectName

        'Set
        _blnImitation = blnImitation

        'Preset
        Select Case _strThisObjectName
            Case "udcCharacter"
                btmCharacter = gbtmCharacterStand(0)
            Case "udcCharacterOne"
                btmCharacter = gbtmCharacterStandRed(0)
            Case "udcCharacterTwo"
                btmCharacter = gbtmCharacterStandBlue(0)
        End Select

        'Set
        intSpotX = intSpawnX
        intSpotY = intSpawnY
        pntCharacter = New Point(intSpotX, intSpotY)

        'Start
        If blnStartAnimation Then
            Start()
        End If

    End Sub

    Public Sub Start(Optional intAnimatingDelay As Integer = 3000)

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
            If thrAnimating.IsAlive Then
                thrAnimating.Abort()
                blnAborted = True

                'While thrAnimating.IsAlive
                '    System.Threading.Thread.Sleep(1)
                '    Debug.Print("CHARACTER STOPANDDISPOSE")
                'End While

                thrAnimating = Nothing
            End If
        End If

    End Sub

    Private Sub Animating()

        'Reset
        blnAborted = False

        'Declare
        Dim intLoop As Integer = 0

        'Continue
        While blnKeepUsingAnimatingThread

            'Debug.Print("CHARACTER")

            'Check for first time pass shooting
            If blnFirstTimeShootingPass Then
                'Default
                pntCharacter.X = intSpotX
                'Set
                blnFirstTimeShootingPass = False
                'Play shot sound
                Dim udcGunShotSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GunShot.mp3", 1000, gintSoundVolume)
                'Change frame immediately
                intFrame = 3
                Select Case _strThisObjectName
                    Case "udcCharacter"
                        btmCharacter = gbtmCharacterShoot(0)
                    Case "udcCharacterOne"
                        btmCharacter = gbtmCharacterShootRed(0)
                    Case "udcCharacterTwo"
                        btmCharacter = gbtmCharacterShootBlue(0)
                End Select
            End If

            'Check for first time pass reloading
            If blnFirstTimeReloadingPass Then
                'Default
                pntCharacter.X = intSpotX
                'Set
                blnFirstTimeReloadingPass = False
                'Play reloading sound
                udcReloadingSound = New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Reloading.mp3", 3000, gintSoundVolume)
                'Change frame immediately
                intFrame = 5
                Select Case _strThisObjectName
                    Case "udcCharacter"
                        btmCharacter = gbtmCharacterReload(0)
                    Case "udcCharacterOne"
                        btmCharacter = gbtmCharacterReloadRed(0)
                    Case "udcCharacterTwo"
                        btmCharacter = gbtmCharacterReloadBlue(0)
                End Select
            End If

            'Check for first time pass running
            If blnFirstTimeRunningPass Then
                'Check incase level is completed
                If blnEndOfLevel Then
                    'Stop running
                    IsRunning = False
                Else
                    'Set
                    blnFirstTimeRunningPass = False
                    'Change frame immediately
                    intFrame = 26
                    'Move point
                    pntCharacter.X = -40
                    'Show
                    Select Case _strThisObjectName
                        Case "udcCharacter"
                            btmCharacter = gbtmCharacterRunning(0)
                    End Select
                End If
            End If

            'Sleep
            System.Threading.Thread.Sleep(_intAnimatingDelay)

            'Check frame
            Select Case intFrame

                Case 1 'Standing, sleep here is 3000
                    'Set frame
                    intFrame = 2
                    Select Case _strThisObjectName
                        Case "udcCharacter"
                            btmCharacter = gbtmCharacterStand(1)
                        Case "udcCharacterOne"
                            btmCharacter = gbtmCharacterStandRed(1)
                        Case "udcCharacterTwo"
                            btmCharacter = gbtmCharacterStandBlue(1)
                    End Select

                Case 2
                    'Set frame
                    intFrame = 1
                    Select Case _strThisObjectName
                        Case "udcCharacter"
                            btmCharacter = gbtmCharacterStand(0)
                        Case "udcCharacterOne"
                            btmCharacter = gbtmCharacterStandRed(0)
                        Case "udcCharacterTwo"
                            btmCharacter = gbtmCharacterStandBlue(0)
                    End Select

                Case 3 'Shooting, sleep here is 250
                    'Set frame
                    intFrame = 4
                    Select Case _strThisObjectName
                        Case "udcCharacter"
                            btmCharacter = gbtmCharacterShoot(1)
                        Case "udcCharacterOne"
                            btmCharacter = gbtmCharacterShootRed(1)
                        Case "udcCharacterTwo"
                            btmCharacter = gbtmCharacterShootBlue(1)
                    End Select

                Case 4 'Neutral to get back to standing
                    'Check for running
                    If blnPrepareToRun Then
                        'Set
                        blnIsRunning = True
                        'Set
                        _intAnimatingDelay = 80
                        'Set
                        blnPrepareToRun = False
                        'Change frame immediately
                        intFrame = 26
                        'Move point
                        pntCharacter.X = -40
                        'Show
                        Select Case _strThisObjectName
                            Case "udcCharacter"
                                btmCharacter = gbtmCharacterRunning(0)
                        End Select
                    Else
                        'Reset if running
                        blnIsRunning = False
                        pntCharacter.X = intSpotX
                        'Check if needs to reload first
                        If blnNeedsToShowReloading Then
                            'Send data
                            SendData()
                            'Set
                            blnIsShooting = False
                            blnIsReloading = True
                            'Set
                            blnNeedsToShowReloading = False
                            'Set
                            _intAnimatingDelay = 100
                            'Play reloading sound
                            udcReloadingSound = New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Reloading.mp3", 3000, gintSoundVolume)
                            'Change frame immediately
                            intFrame = 5
                            Select Case _strThisObjectName
                                Case "udcCharacter"
                                    btmCharacter = gbtmCharacterReload(0)
                                Case "udcCharacterOne"
                                    btmCharacter = gbtmCharacterReloadRed(0)
                                Case "udcCharacterTwo"
                                    btmCharacter = gbtmCharacterReloadBlue(0)
                            End Select
                            'Increase
                            intReloadTimes += 1
                        Else
                            'Standing
                            StandingFrame()
                        End If
                    End If

                Case 5 To 24 'Reloading, sleep here is 100
                    'Set frame
                    intFrame += 1
                    Select Case _strThisObjectName
                        Case "udcCharacter"
                            btmCharacter = gbtmCharacterReload(intFrame - 5)
                        Case "udcCharacterOne"
                            btmCharacter = gbtmCharacterReloadRed(intFrame - 5)
                        Case "udcCharacterTwo"
                            btmCharacter = gbtmCharacterReloadBlue(intFrame - 5)
                    End Select

                Case 25
                    'Set frame
                    intFrame = 4 'Goes back to neutral standing
                    Select Case _strThisObjectName
                        Case "udcCharacter"
                            btmCharacter = gbtmCharacterReload(21)
                        Case "udcCharacterOne"
                            btmCharacter = gbtmCharacterReloadRed(21)
                        Case "udcCharacterTwo"
                            btmCharacter = gbtmCharacterReloadBlue(21)
                    End Select
                    'Reset bullets
                    _intBullets = 0

                Case 26 To 41
                    'Check if end of level
                    If blnEndOfLevel Then
                        'Stop running
                        IsRunning = False
                    Else
                        'Set frame
                        intFrame += 1
                        'Move point
                        pntCharacter.X = -40
                        Select Case _strThisObjectName
                            Case "udcCharacter"
                                btmCharacter = gbtmCharacterRunning(intFrame - 26) '27 - 26 = 1 in the array
                        End Select
                        'Play reloading sound
                        Select Case intFrame
                            Case 32, 39, 41
                                Dim udcStepSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Step.mp3", 1250, gintSoundVolume)
                        End Select
                        'Check if stop running
                        If intFrame = 42 Then
                            intFrame = 4
                        End If
                    End If

            End Select

        End While

    End Sub

    Private Sub StandingFrame()

        'Change sleep
        _intAnimatingDelay = 3000

        'Set
        blnIsShooting = False
        blnIsReloading = False

        'Set frame
        intFrame = 1

        'Set
        Select Case _strThisObjectName
            Case "udcCharacter"
                btmCharacter = gbtmCharacterStand(0)
            Case "udcCharacterOne"
                btmCharacter = gbtmCharacterStandRed(0)
            Case "udcCharacterTwo"
                btmCharacter = gbtmCharacterStandBlue(0)
        End Select

    End Sub

    Private Sub SendData()

        'Send data
        If blnSendData Then
            'Send
            gSendData("5|")
            'Set
            blnSendData = False
        End If

    End Sub

    Public Sub Shot()

        'Check for instance
        If thrAnimating IsNot Nothing Then
            'Set
            blnIsRunning = False
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
                    blnNeedsToShowReloading = True
                End If
            End If
            'Abort thread
            thrAnimating.Abort()
            blnAborted = True

            'While thrAnimating.IsAlive
            '    System.Threading.Thread.Sleep(1)
            '    Debug.Print("CHARACTER SHOT")
            'End While

            'Set
            blnIsShooting = True
            'Set
            blnFirstTimeShootingPass = True
            'Restart thread
            Start(250)
        End If

    End Sub

    Public Sub Reload()

        'Check if shooting
        If blnIsShooting Then
            'Set
            blnNeedsToShowReloading = True
        Else
            'Set
            blnIsRunning = False
            'Check for instance
            If thrAnimating IsNot Nothing Then
                'Send data
                SendData()
                'Abort thread
                thrAnimating.Abort()
                blnAborted = True

                'While thrAnimating.IsAlive
                '    System.Threading.Thread.Sleep(1)
                '    Debug.Print("CHARACTER RELOAD")
                'End While

                'Set
                blnIsReloading = True
                'Set
                blnFirstTimeReloadingPass = True
                'Increase
                intReloadTimes += 1
                'Restart thread
                Start(100)
            End If
        End If

    End Sub

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

    Public ReadOnly Property IsReloading() As Boolean

        'Return
        Get
            Return blnIsReloading
        End Get

    End Property

    Public ReadOnly Property IsShooting() As Boolean

        'Return
        Get
            Return blnIsShooting
        End Get

    End Property

    Public Property IsRunning() As Boolean

        'Return
        Get
            Return blnIsRunning
        End Get

        'Set
        Set(value As Boolean)

            'Set
            blnIsRunning = value

            'Check
            If Not value Then
                'Check for instance
                If thrAnimating IsNot Nothing Then
                    'Abort
                    While thrAnimating.IsAlive
                        If Not blnAborted Then
                            thrAnimating.Abort() 'If a thread is trying to abort multiple times at the exact same time, it does affect processor speed, and creates crazy glitches
                            blnAborted = True
                        End If
                    End While

                    'While thrAnimating.IsAlive
                    '    System.Threading.Thread.Sleep(1)
                    '    Debug.Print("CHARACTER ISRUNNING")
                    'End While

                    'Set
                    pntCharacter.X = intSpotX
                    'Stand
                    StandingFrame()
                    'Start
                    Start()
                End If
            End If

        End Set

    End Property

    Public Property EndOfLevel() As Boolean

        'Return
        Get
            Return blnEndOfLevel
        End Get

        'Set
        Set(value As Boolean)

            'Set
            blnEndOfLevel = value

            'Check if ending level
            If value Then
                blnKeepUsingAnimatingThread = False
            End If

        End Set

    End Property

    Public WriteOnly Property PrepareToRun() As Boolean

        'Return
        Set(value As Boolean)
            blnPrepareToRun = value
        End Set

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
        End If

    End Sub

    Public Sub Running()

        'Check for instance
        If thrAnimating IsNot Nothing Then
            'Abort
            While thrAnimating.IsAlive
                If Not blnAborted Then
                    thrAnimating.Abort() 'If a thread is trying to abort multiple times at the exact same time, it does affect processor speed, and creates crazy glitches
                    blnAborted = True
                End If
            End While

            'While thrAnimating.IsAlive
            '    System.Threading.Thread.Sleep(1)
            '    Debug.Print("CHARACTER RUNNING")
            'End While

            'Set
            blnIsRunning = True
            'Set
            blnFirstTimeRunningPass = True
            'Restart thread
            Start(80)
        End If

    End Sub

End Class