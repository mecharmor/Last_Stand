'Options
Option Explicit On
Option Strict On
Option Infer Off

Module mdlGlobal

    'Gameplay speed
    Public Const gintCHARACTER_MOVEMENT_SPEED As Integer = 27

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

    'Level
    Public gintLevel As Integer = 1

    'Character memory for class file
    Public gabtmCharacterStandMemories(1) As Bitmap
    Public gabtmCharacterShootMemories(1) As Bitmap
    Public gabtmCharacterReloadMemories(21) As Bitmap
    Public gabtmCharacterRunningMemories(16) As Bitmap
    Public gabtmCharacterStandRedMemories(1) As Bitmap
    Public gabtmCharacterShootRedMemories(1) As Bitmap
    Public gabtmCharacterReloadRedMemories(21) As Bitmap
    Public gabtmCharacterStandBlueMemories(1) As Bitmap
    Public gabtmCharacterShootBlueMemories(1) As Bitmap
    Public gabtmCharacterReloadBlueMemories(21) As Bitmap

    'Zombies memory for class file
    Public gabtmZombieWalkMemories(3) As Bitmap
    Public gabtmZombieDeath1Memories(5) As Bitmap
    Public gabtmZombieDeath2Memories(5) As Bitmap
    Public gabtmZombiePinMemories(1) As Bitmap
    Public gabtmZombieWalkRedMemories(3) As Bitmap
    Public gabtmZombieDeathRed1Memories(5) As Bitmap
    Public gabtmZombieDeathRed2Memories(5) As Bitmap
    Public gabtmZombiePinRedMemories(1) As Bitmap
    Public gabtmZombieWalkBlueMemories(3) As Bitmap
    Public gabtmZombieDeathBlue1Memories(5) As Bitmap
    Public gabtmZombieDeathBlue2Memories(5) As Bitmap
    Public gabtmZombiePinBlueMemories(1) As Bitmap

    'Chained zombie memory for class file
    Public gabtmChainedZombieMemories(2) As Bitmap

    'Zombie face
    Public gabtmFaceZombieMemories(1) As Bitmap

    'Helicopter
    Public gabtmHelicopterMemories(4) As Bitmap

    'Zombies class, must be public because used in multiple threads of other classes, game movement with zombies moving
    Public gaudcZombies() As clsZombie
    Public gaudcZombiesOne() As clsZombie
    Public gaudcZombiesTwo() As clsZombie

    'Chained zombie class, must be public because used in multiple threads of other classes, game movement with zombies moving
    Public gudcChainedZombie As clsChainedZombie

    'Face zombie
    Public gudcFaceZombie As clsFaceZombie

    'Pipe valve
    Public gpntPipeValve As New Point(0, 0)

    'Multiplayer versus
    Public gswClientData As IO.StreamWriter

    'Sending data cleared
    Public gblnSendingDataCleared As Boolean = False
    Public gblnReceivingDataCleared As Boolean = False

    'Data stream buffer
    Public gstrData As String = ""

    Public Sub gSendData(intDataCase As Integer, strData As String)

        'Wait here
        While Not gblnSendingDataCleared
        End While

        'Set
        gblnSendingDataCleared = False

        'Add delimiter to the string
        strData = CStr(intDataCase) & "|" & strData & "~"

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

        'Set
        gblnSendingDataCleared = True

    End Sub

    Public Sub gMoveZombiesWhileRunning(audcZombiesType() As clsZombie)

        'Move zombies
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            'Check if spawned
            If audcZombiesType(intLoop).Spawned Then
                'Set new point so that the zombie is moving
                audcZombiesType(intLoop).Point = New Point(audcZombiesType(intLoop).Point.X - gintCHARACTER_MOVEMENT_SPEED, audcZombiesType(intLoop).Point.Y)
            End If
        Next

        'Check level
        Select Case gintLevel
            Case 2
                'Set new point for pipe
                gpntPipeValve.X = gpntPipeValve.X - CInt(gintCHARACTER_MOVEMENT_SPEED * 0.75)
                'Set new point for face zombie
                gudcFaceZombie.Point = New Point(gudcFaceZombie.Point.X - CInt(gintCHARACTER_MOVEMENT_SPEED * 0.75), gudcChainedZombie.Point.Y)
                'Set new point so that the zombie is moving
                gudcChainedZombie.Point = New Point(gudcChainedZombie.Point.X - gintCHARACTER_MOVEMENT_SPEED, gudcChainedZombie.Point.Y)
        End Select

    End Sub

End Module