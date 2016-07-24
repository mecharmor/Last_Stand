'Options
Option Explicit On
Option Strict On
Option Infer Off

Module mdlGlobal

    'Gameplay speed
    Public Const gCHARACTER_MOVEMENT_SPEED As Integer = 55

    'Screen width ratio
    Public gdblScreenWidthRatio As Double = 0
    Public gdblScreenHeightRatio As Double = 0

    'Used for sound
    Public gintAlias As Integer = 0
    Public gintSoundVolume As Integer = 1000

    'Game play variables
    Public gintBullets As Integer = 0 'Used for key press event, to make key press threading match with rendering thread

    'Game movement
    Public gpntGameBackground As New Point(0, 0) 'Used in the character class and form game

    'Character
    Public gbtmCharacterStand(1) As Bitmap
    Public gbtmCharacterShoot(1) As Bitmap
    Public gbtmCharacterReload(21) As Bitmap
    Public gbtmCharacterRunning(16) As Bitmap
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
    Public gbtmZombieWalkRed(3) As Bitmap
    Public gbtmZombieDeathRed1(5) As Bitmap
    Public gbtmZombieDeathRed2(5) As Bitmap
    Public gbtmZombiePinRed(1) As Bitmap
    Public gbtmZombieWalkBlue(3) As Bitmap
    Public gbtmZombieDeathBlue1(5) As Bitmap
    Public gbtmZombieDeathBlue2(5) As Bitmap
    Public gbtmZombiePinBlue(1) As Bitmap

    'Zombies class, must be public because used in multiple threads of other classes
    Public gaudcZombies() As clsZombie
    Public gaudcZombiesOne() As clsZombie
    Public gaudcZombiesTwo() As clsZombie

    'Zombies that produce game lose
    Public gbtmUnderwaterZombieGameLose As Bitmap

    'Helicopter
    Public gbtmHelicopter(4) As Bitmap

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

    Public Sub gMoveZombiesWhileRunning(audcZombiesType() As clsZombie)

        'Move zombies
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            'Check if spawned
            If audcZombiesType(intLoop).Spawned Then
                'Check if not marked to die
                If Not audcZombiesType(intLoop).MarkedToDie Then
                    audcZombiesType(intLoop).ZombiePoint = New Point(audcZombiesType(intLoop).ZombiePoint.X - gCHARACTER_MOVEMENT_SPEED,
                                                           audcZombiesType(intLoop).ZombiePoint.Y)
                End If
            End If
        Next

    End Sub

End Module