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
    Private blnDontSendData As Boolean = False

    'Reload sound
    Private udcReloadingSound As clsSound

    'Reload count
    Private intReloadTimes As Integer = 0

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, strThisObjectName As String, Optional blnStart As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set
        _strThisObjectName = strThisObjectName

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
        If blnStart Then
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
                thrAnimating = Nothing
            End If
        End If

    End Sub

    Private Sub Animating()

        'Declare
        Dim intLoop As Integer = 0

        'Continue
        While True

            'Check for first time pass shooting
            If blnFirstTimeShootingPass Then
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
                    'Check if needs to reload first
                    If blnNeedsToShowReloading Then
                        'Send data
                        Select Case _strThisObjectName
                            Case "udcCharacterOne", "udcCharacterTwo"
                                If blnDontSendData Then
                                    blnDontSendData = False
                                Else
                                    gSendData("4|")
                                End If
                        End Select
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
                        'Change sleep
                        _intAnimatingDelay = 3000
                        'Set
                        blnIsShooting = False
                        blnIsReloading = False
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
                    Select Case _strThisObjectName
                        Case "udcCharacter"
                            gintBullet = 0
                        Case "udcCharacterOne"
                            gintBulletOne = 0
                        Case "udcCharacterTwo"
                            gintBulletTwo = 0
                    End Select

            End Select

        End While

    End Sub

    Public Sub Shot()

        'Check for instance
        If thrAnimating IsNot Nothing Then
            'Increase
            Select Case _strThisObjectName
                Case "udcCharacter"
                    gintBullet += 1
                    'Set
                    If gintBullet >= 30 Then
                        blnNeedsToShowReloading = True
                    End If
                Case "udcCharacterOne"
                    gintBulletOne += 1
                    'Set
                    If gintBulletOne >= 30 Then
                        blnNeedsToShowReloading = True
                    End If
                Case "udcCharacterTwo"
                    gintBulletTwo += 1
                    'Set
                    If gintBulletTwo >= 30 Then
                        blnNeedsToShowReloading = True
                    End If
            End Select
            'Abort thread
            thrAnimating.Abort()
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
            'Check for instance
            If thrAnimating IsNot Nothing Then
                'Send data
                Select Case _strThisObjectName
                    Case "udcCharacterOne", "udcCharacterTwo"
                        If blnDontSendData Then
                            blnDontSendData = False
                        Else
                            gSendData("4|")
                        End If
                End Select
                'Abort thread
                thrAnimating.Abort()
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

    Public WriteOnly Property DontSendData() As Boolean

        'Set
        Set(value As Boolean)
            blnDontSendData = value
        End Set

    End Property

    Public Sub StopReloadingSound()

        'Stop sound
        If udcReloadingSound IsNot Nothing Then
            udcReloadingSound.StopAndCloseSound()
        End If

    End Sub

End Class