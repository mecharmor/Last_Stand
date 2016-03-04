'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Notes: Original menu pixel size is 1920x1200px

    'Shadows Event Paint(ByVal g As Graphics)

    'Declare main
    Private intScreenWidth As Integer = 0
    Private intScreenHeight As Integer = 0
    Private picMenu As PictureBox
    Private picStart As PictureBox
    Private picBackground As PictureBox
    Private picZombie As PictureBox
    Private thrZombie As System.Threading.Thread
    Private intWalkingCounter As Integer = 0
    Private blnZombieSwitch As Boolean = False

    Private Sub frmGame_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Screen
        intScreenWidth = Me.Width '1696
        intScreenHeight = Me.Height '1066

        'Add picture box menu
        picMenu = New PictureBox()
        Dim btmMenu As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Menu.jpg"), intScreenWidth, intScreenHeight)
        picMenu.Image = btmMenu
        picMenu.Width = intScreenWidth
        picMenu.Height = intScreenHeight
        picMenu.Left = 0
        picMenu.Top = 0
        Controls.Add(picMenu)

        'Add picture box start
        picStart = New PictureBox()
        picStart.Visible = False
        Dim btmStart As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Start.png"), 250, 80)
        picStart.Image = btmStart
        picStart.Width = 250
        picStart.Height = 80
        picStart.Left = 1100
        picStart.Top = 50
        Controls.Add(picStart)
        picStart.BringToFront()
        picStart.BackColor = Color.Transparent
        picStart.Parent = picMenu
        picStart.Visible = True

        'Add event
        AddHandler picStart.Click, AddressOf picStart_Click

    End Sub

    Private Sub frmGame_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Stop thread
        If thrZombie.IsAlive Then
            thrZombie.Abort()
        End If

    End Sub

    Private Sub picStart_Click(sender As Object, e As EventArgs)

        'Remove event
        RemoveHandler picStart.Click, AddressOf picStart_Click

        'Add picture box background
        picBackground = New PictureBox()
        Dim btmBackground As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Background.jpg"), intScreenWidth, intScreenHeight)
        picBackground.Image = btmBackground
        picBackground.Width = intScreenWidth
        picBackground.Height = intScreenHeight
        picBackground.Left = 0
        picBackground.Top = 0
        Controls.Add(picBackground)

        'Start a zombie on the right
        picZombie = New PictureBox()
        Dim btmZombie As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/1.png"), 1500, 1500)
        picZombie.Image = btmZombie
        picZombie.Width = 1500
        picZombie.Height = 1500
        picZombie.Left = Me.Width + 100
        picZombie.Top = -100
        Controls.Add(picZombie)
        picZombie.BringToFront()
        picZombie.BackColor = Color.Transparent
        picZombie.Parent = picBackground

        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, True)

        'Start a moving thread, this is sort of like a timer but we are using a thread from our CPU
        thrZombie = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf ZombieWalking))
        thrZombie.Start()

        'Remove the previous picture boxes
        picMenu.Dispose()
        picMenu = Nothing
        picStart.Dispose()
        picStart = Nothing

    End Sub

    Private Sub ZombieWalking()

        'Notes: http://www.vbdotnetforums.com/gui/15271-picture-box-refresh-problem.html
        '       http://www.codeproject.com/Articles/409988/Beginners-Starting-a-D-Game-with-GDIplus
        '       The refresh problem with the picture box is normal, this happens. Above links talk about not using picture boxes. I guess we have to figure out GDI+.
        '       Zombie frames = 1, 2, 3, 4, 3, 2, 1, 2, 3, 4

        'Keep moving
        While True
            'Reduce CPU usage, or lag the zombie on purpose
            System.Threading.Thread.Sleep(15) 'Speed of miliseconds in sleeping, waiting
            Debug.Print(CStr(intWalkingCounter))
            'Set and increase
            If Not blnZombieSwitch Then
                If intWalkingCounter = 4 Then
                    blnZombieSwitch = True
                    intWalkingCounter -= 1
                Else
                    intWalkingCounter += 1
                End If
            Else
                If intWalkingCounter = 1 Then
                    blnZombieSwitch = False
                    intWalkingCounter += 1
                Else
                    intWalkingCounter -= 1
                End If
            End If
            'Image swap
            Dim btmZombie As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/" & CStr(intWalkingCounter) & ".png"), 1500, 1500)
            picZombie.Invoke(Sub() picZombie.Image = btmZombie) 'Controls must invoke first before doing what you want, this is inside a thread
            'Move zombie
            picZombie.Invoke(Sub() picZombie.Left -= 35) 'Controls must invoke first before doing what you want, this is inside a thread
        End While

    End Sub

End Class