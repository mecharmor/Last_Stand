'Options
Option Explicit On
Option Strict On
Option Infer Off

Module mdlGlobal

    'Gameplay speed
    Public Const gintCHARACTER_MOVEMENT_SPEED As Integer = 27

    'Screen width ratio
    Public gsngScreenWidthRatio As Single = 0F
    Public gsngScreenHeightRatio As Single = 0F

    'Used for sound
    Public gintAlias As Integer = 0
    Public gintSoundVolume As Integer = 1000

    'Game movement
    Public gpntGameBackground As New Point(0, 0) 'Used in the character class and form game

    'Stop watch for typing
    Public gswhTimeTyped As Stopwatch

    'Stop meaning game is over either by pin or made it to an escape
    Public gblnPreventKeyPressEvent As Boolean = False

    'Level
    Public gintLevel As Integer = 1

    'On screen words
    Public gabtmOnScreenWordMissedRedMemories(1) As Bitmap
    Public gabtmOnScreenWordMissedBlueMemories(1) As Bitmap

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

    'Destroyed brick wall
    Public gpntDestroyedBrickWall As New Point(0, 0)

    'Pipe valve
    Public gpntPipeValve As New Point(0, 0)

    'Face zombie
    Public gudcFaceZombie As clsFaceZombie

    'Multiplayer versus
    Public gswClientData As IO.StreamWriter

    'Data stream buffer
    Public gstrData As String = ""

    Public Function gGetRandomNumber(intBeginningNumber As Integer, intEndingNumber As Integer) As Integer

        'Declare
        Static srndNumber As New Random 'Use static to make the random very random

        'Return
        Return srndNumber.Next(intBeginningNumber, intEndingNumber + 1)

    End Function

    Public Sub gCloseApplicationWithErrorMessage(strErrorMessage As String)

        'Display
        MessageBox.Show(strErrorMessage, "Last Stand", MessageBoxButtons.OK, MessageBoxIcon.Error)

        'Exit the program
        End

    End Sub

    Public Sub gSendData(intDataCase As Integer, Optional strData As String = "")

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

    End Sub

    Public Sub gMoveZombiesWhileRunning(audcZombiesType() As clsZombie)

        'Move zombies
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            'Check if spawned
            If audcZombiesType(intLoop).Spawned Then
                'Set new point so that the zombie is moving
                audcZombiesType(intLoop).Point = New Point(audcZombiesType(intLoop).Point.X - gintCHARACTER_MOVEMENT_SPEED,
                                                 audcZombiesType(intLoop).Point.Y)
            End If
        Next

        'Check level
        Select Case gintLevel
            Case 1
                'Set new point for destroyed brick wall
                gpntDestroyedBrickWall.X = gpntDestroyedBrickWall.X - CInt(gintCHARACTER_MOVEMENT_SPEED * 0.9)
            Case 2
                'Set new point for pipe
                gpntPipeValve.X = gpntPipeValve.X - CInt(gintCHARACTER_MOVEMENT_SPEED * 0.75)
                'Set new point for face zombie
                gudcFaceZombie.Point = New Point(gudcFaceZombie.Point.X - CInt(gintCHARACTER_MOVEMENT_SPEED * 0.75), gudcChainedZombie.Point.Y)
                'Set new point so that the zombie is moving
                gudcChainedZombie.Point = New Point(gudcChainedZombie.Point.X - gintCHARACTER_MOVEMENT_SPEED, gudcChainedZombie.Point.Y)
        End Select

    End Sub

    Public Sub gStopAndDisposeTimerVariable(tmrTimer As Timers.Timer)

        'Disable timers
        tmrTimer.Enabled = False

        'Stop and dispose timers
        tmrTimer.Stop()
        tmrTimer.Dispose()

    End Sub

End Module