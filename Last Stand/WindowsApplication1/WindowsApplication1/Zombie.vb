'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class Zombie

    'Declare
    Private intFrame As Integer = 1
    Private thrAnimation As System.Threading.Thread
    Private blnSwitch As Boolean = False
    Private intMovementPoint As Integer = 0
    Private _intSpeed As Integer = 25

    'Bitmaps
    Private btmZombieWalk1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/1.png"), 765, 752)
    Private btmZombieWalk2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/2.png"), 765, 752)
    Private btmZombieWalk3 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/3.png"), 765, 752)
    Private btmZombieWalk4 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/4.png"), 765, 752)
    Private btmZombieDeath1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/1.png"), 765, 752)
    Private btmZombieDeath2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/2.png"), 765, 752)
    Private btmZombieDeath3 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/3.png"), 765, 752)
    Private btmZombieDeath4 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/4.png"), 765, 752)
    Private btmZombieDeath5 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/5.png"), 765, 752)
    Private btmZombieDeath6 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/6.png"), 765, 752)
    Public btmZombie As Bitmap
    Public rectZombie As Rectangle

    'Sound
    Private udcDeathSound As Sound

    'Death
    Private blnAlive As Boolean = True

    Public Sub New(intSpawnPosition As Integer, Optional blnStart As Boolean = False, Optional intSpeed As Integer = 25)

        'Set
        _intSpeed = intSpeed

        'Set
        intMovementPoint = intSpawnPosition

        'Set sound
        udcDeathSound = New Sound("ZombieDeath1", AppDomain.CurrentDomain.BaseDirectory & "/Sounds/ZombieDeath1.mp3")

        'Set thread
        thrAnimation = New System.Threading.Thread(Sub() Animating())

        'Start
        If blnStart Then
            Start()
        End If

    End Sub

    Public Sub Start()

        'Start thread
        thrAnimation.Start()

    End Sub

    Private Sub Animating()

        'Loop
        While True
            'Frame
            If blnAlive Then
                'Change rectangle
                rectZombie = New Rectangle(intMovementPoint, 400, 765, 752)
                intMovementPoint -= _intSpeed 'Speed they come at
                'Check walking
                Select Case intFrame
                    Case 1
                        btmZombie = btmZombieWalk1
                    Case 2
                        btmZombie = btmZombieWalk2
                    Case 3
                        btmZombie = btmZombieWalk3
                    Case 4
                        btmZombie = btmZombieWalk4
                End Select
                'Reduce CPU usage
                System.Threading.Thread.Sleep(175) '250
            Else
                'Check death
                Select Case intFrame
                    Case 1
                        btmZombie = btmZombieDeath1
                    Case 2
                        btmZombie = btmZombieDeath2
                    Case 3
                        btmZombie = btmZombieDeath3
                    Case 4
                        btmZombie = btmZombieDeath4
                    Case 5
                        btmZombie = btmZombieDeath5
                    Case 6
                        btmZombie = btmZombieDeath6
                End Select
                'Reduce CPU usage
                System.Threading.Thread.Sleep(50)
            End If
            'Update the frame
            If blnAlive Then
                If Not blnSwitch Then
                    intFrame += 1
                    If intFrame = 5 Then
                        blnSwitch = True
                        intFrame = 4
                    End If
                Else
                    intFrame -= 1
                    If intFrame = 0 Then
                        blnSwitch = False
                        intFrame = 1
                    End If
                End If
            Else
                intFrame += 1
                If intFrame = 7 Then
                    'Zombie is dead, can remove from game
                    thrAnimation.Abort()
                End If
            End If
        End While

    End Sub

    Public ReadOnly Property IsAlive() As Boolean

        'Return alive or not
        Get
            Return blnAlive
        End Get

    End Property

    Public Sub Dead()

        'Change
        intFrame = 1

        'Set
        blnAlive = False

        'Play sound of death
        udcDeathSound.PlaySound(False)

    End Sub

    Public Sub StopZombie()

        'Abort thread
        If thrAnimation.IsAlive Then
            thrAnimation.Abort()
        End If

    End Sub

End Class