'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsHelicopter

    'Constant
    Private Const dblHELICOPTER_ROTATINGBLADE_DELAY As Double = 120

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 0

    'Sound
    Private _udcRotatingBladeSound As clsSound

    'Bitmaps
    Private btmImage As Bitmap
    Private pntPoint As Point

    'Timer
    Private tmrAnimation As New System.Timers.Timer

    Public Sub New(frmToPass As Form, udcRotatingBladeSound As clsSound, Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set animation
        btmImage = gabtmHelicopterMemories(0)

        'Set
        pntPoint = New Point(1439, 0)

        'Set sound
        _udcRotatingBladeSound = udcRotatingBladeSound

        'Start the sound in a loop
        _udcRotatingBladeSound.PlaySound(gintSoundVolume, True)

        'Set timer
        tmrAnimation.AutoReset = True

        'Add handlers
        AddHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

        'Start
        If blnStartAnimation Then
            Start()
        End If

    End Sub

    Private Sub ElapsedAnimation(sender As Object, e As EventArgs)

        'Check the frame
        Select Case intFrame
            Case 0
                btmImage = gabtmHelicopterMemories(1)
            Case 1
                btmImage = gabtmHelicopterMemories(2)
            Case 2
                btmImage = gabtmHelicopterMemories(3)
            Case 3
                btmImage = gabtmHelicopterMemories(4)
            Case 4
                btmImage = gabtmHelicopterMemories(0)
        End Select

        'Increase frame
        intFrame += 1

        'Check frame
        If intFrame = 5 Then
            intFrame = 0
        End If

    End Sub


    Public Sub Start(Optional dblAnimatingDelay As Double = dblHELICOPTER_ROTATINGBLADE_DELAY)

        'Set timer delay
        tmrAnimation.Interval = dblHELICOPTER_ROTATINGBLADE_DELAY

        'Start the animating thread
        tmrAnimation.Enabled = True

    End Sub

    Public Sub StopAndDispose()

        'Disable timers
        tmrAnimation.Enabled = False

        'Stop and dispose timers
        tmrAnimation.Stop()
        tmrAnimation.Dispose()

        'Remove handlers
        RemoveHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

    End Sub

    Public Property Point() As Point

        'Return
        Get
            Return pntPoint
        End Get

        'Set
        Set(value As Point)
            pntPoint = value
        End Set

    End Property

    Public ReadOnly Property Image() As Bitmap

        'Return
        Get
            Return btmImage
        End Get

    End Property

End Class