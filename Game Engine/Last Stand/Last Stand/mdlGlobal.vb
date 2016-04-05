'Options
Option Explicit On
Option Strict On
Option Infer Off

Module mdlGlobal

    'Screen width ratio
    Public gdblScreenWidthRatio As Double = 0
    Public gdblScreenHeightRatio As Double = 0

    'Used for sound
    Public gintAlias As Integer = 0
    Public gintSoundVolume As Integer = 1000

    'Used for reloading
    Public gintBullet As Integer = 0
    Public gintBulletOne As Integer = 0
    Public gintBulletTwo As Integer = 0

    'Character
    Public gbtmCharacterStand(1) As Bitmap
    Public gbtmCharacterShoot(1) As Bitmap
    Public gbtmCharacterReload(21) As Bitmap
    Public gbtmCharacterStandRed(1) As Bitmap
    Public gbtmCharacterShootRed(1) As Bitmap
    Public gbtmCharacterReloadRed(21) As Bitmap
    Public gbtmCharacterStandBlue(1) As Bitmap
    Public gbtmCharacterShootBlue(1) As Bitmap
    Public gbtmCharacterReloadBlue(21) As Bitmap

    'Zombies
    Public gbtmZombieWalk(3) As Bitmap
    Public gbtmZombieDeath1(5) As Bitmap
    Public gbtmZombieDeath2(5) As Bitmap
    Public gbtmZombiePin(1) As Bitmap

    'Multiplayer versus
    Public gswClientData As IO.StreamWriter

    Public Sub gSendData(strData As String)

        'Send data
        Try
            'Send
            gswClientData.WriteLine(strData)
            gswClientData.Flush()
            'Wait
            Application.DoEvents()
        Catch ex As Exception
            'No debug
        End Try

    End Sub

End Module