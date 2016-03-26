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

    'Bitmaps
    Public btmCharacter As Bitmap
    Public pntCharacter As Point

    'Thread
    Private thrAnimating As System.Threading.Thread
    Private _intAnimatingDelay As Integer = 0
    Private blnCharacterIsShooting As Boolean = False
    Private blnCharacterIsReloading As Boolean = False
    Private blnNeedsToShowReloading As Boolean = False
    Private blnLoop As Boolean = True

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, Optional blnStart As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Preset
        btmCharacter = gbtmCharacterStand(0)

        'Set
        intSpotX = intSpawnX
        intSpotY = intSpawnY
        pntCharacter = New Point(intSpotX, intSpotY)

        'Start
        If blnStart Then
            Start(3000)
        End If

    End Sub

    Public Sub Start(intAnimatingDelay As Integer)

        'Set
        _intAnimatingDelay = intAnimatingDelay

        'Start thread
        thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
        thrAnimating.Start()

    End Sub

    Public Sub StopAndDispose()

        'Abort
        thrAnimating.Abort()

        'Dispose
        thrAnimating = Nothing

    End Sub

    Private Sub AnimatingSleeping()

        'Declare
        Dim intLoop As Integer = 0

        'Sleep
        Do Until intLoop = _intAnimatingDelay Or Not blnLoop
            System.Threading.Thread.Sleep(1)
            intLoop += 1
        Loop

    End Sub

    Private Sub Animating()

        'Continue
        While True

            'Sleep
            AnimatingSleeping()

            'Exit
            If Not blnLoop Then
                'Shooting
                If blnCharacterIsShooting Then
                    'Change frame immediately
                    intFrame = 3
                    btmCharacter = gbtmCharacterShoot(0)
                    'Sleep
                    _intAnimatingDelay = 250
                End If
                'Reloading
                If blnCharacterIsReloading Then
                    'Change frame immediately
                    intFrame = 5
                    btmCharacter = gbtmCharacterReload(0)
                    'Sleep
                    _intAnimatingDelay = 100
                End If
            End If

            'Check to restart
            If Not blnLoop Then
                blnLoop = True
            Else

                'Check frame
                Select Case intFrame

                    Case 1 'Standing, sleep here is 3000
                        'Set frame
                        intFrame = 2
                        btmCharacter = gbtmCharacterStand(1)
                    Case 2
                        'Set frame
                        intFrame = 1
                        btmCharacter = gbtmCharacterStand(0)

                    Case 3 'Shooting, sleep here is 250
                        'Set frame
                        intFrame = 4
                        btmCharacter = gbtmCharacterShoot(1)

                    Case 4 'Neutral to get back to standing
                        'Check if needs to reload first
                        If blnNeedsToShowReloading Then
                            blnNeedsToShowReloading = False
                            CharacterReload()
                        Else
                            'Set
                            blnCharacterIsShooting = False
                            blnCharacterIsReloading = False
                            'Set frame
                            intFrame = 1
                            btmCharacter = gbtmCharacterStand(0)
                            'Change sleep
                            _intAnimatingDelay = 3000
                        End If

                    Case 5 'Reloading, sleep here is 100
                        'Set frame
                        intFrame = 6
                        btmCharacter = gbtmCharacterReload(1)
                    Case 6
                        'Set frame
                        intFrame = 7
                        btmCharacter = gbtmCharacterReload(2)
                    Case 7
                        'Set frame
                        intFrame = 8
                        btmCharacter = gbtmCharacterReload(3)
                    Case 8
                        'Set frame
                        intFrame = 9
                        btmCharacter = gbtmCharacterReload(4)
                    Case 9
                        'Set frame
                        intFrame = 10
                        btmCharacter = gbtmCharacterReload(5)
                    Case 10
                        'Set frame
                        intFrame = 11
                        btmCharacter = gbtmCharacterReload(6)
                    Case 11
                        'Set frame
                        intFrame = 12
                        btmCharacter = gbtmCharacterReload(7)
                    Case 12
                        'Set frame
                        intFrame = 13
                        btmCharacter = gbtmCharacterReload(8)
                    Case 13
                        'Set frame
                        intFrame = 14
                        btmCharacter = gbtmCharacterReload(9)
                    Case 14
                        'Set frame
                        intFrame = 15
                        btmCharacter = gbtmCharacterReload(10)
                    Case 15
                        'Set frame
                        intFrame = 16
                        btmCharacter = gbtmCharacterReload(11)
                    Case 16
                        'Set frame
                        intFrame = 17
                        btmCharacter = gbtmCharacterReload(12)
                    Case 17
                        'Set frame
                        intFrame = 18
                        btmCharacter = gbtmCharacterReload(13)
                    Case 18
                        'Set frame
                        intFrame = 19
                        btmCharacter = gbtmCharacterReload(14)
                    Case 19
                        'Set frame
                        intFrame = 20
                        btmCharacter = gbtmCharacterReload(15)
                    Case 20
                        'Set frame
                        intFrame = 21
                        btmCharacter = gbtmCharacterReload(16)
                    Case 21
                        'Set frame
                        intFrame = 22
                        btmCharacter = gbtmCharacterReload(17)
                    Case 22
                        'Set frame
                        intFrame = 23
                        btmCharacter = gbtmCharacterReload(18)
                    Case 23
                        'Set frame
                        intFrame = 24
                        btmCharacter = gbtmCharacterReload(19)
                    Case 24
                        'Set frame
                        intFrame = 25
                        btmCharacter = gbtmCharacterReload(20)
                    Case 25
                        'Set frame
                        intFrame = 4 'Goes back to neutral standing
                        btmCharacter = gbtmCharacterReload(21)

                End Select

            End If

        End While

    End Sub

    Public Sub CharacterShot(intBullet As Integer)

        'Set
        blnCharacterIsShooting = True

        'Set
        If intBullet = 30 Then
            blnNeedsToShowReloading = True
        End If

        'Play shot sound
        Dim udcGunShotSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GunShot.mp3", 1000, gintSoundVolume)

        'Stop thread
        blnLoop = False

    End Sub

    Public Sub CharacterReload()

        'Set
        blnCharacterIsReloading = True

        'Play reload sound
        Dim udcReloadSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Reload.mp3", 6000, gintSoundVolume)

        'Stop thread
        blnLoop = False

    End Sub

    Public ReadOnly Property CharacterIsReloading() As Boolean

        'Return
        Get
            Return blnCharacterIsReloading
        End Get

    End Property

End Class