'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsZombieGameLose

    'Declare
    Private _frmToPass As Form
    Private _strTypeOfZombie As String = ""
    Private intFrame As Integer = 1

    'For error preventing
    Private blnKeepUsingAnimatingThread As Boolean = True

    'Bitmaps
    Private btmZombie As Bitmap
    Private pntZombie As Point

    'Thread
    Private thrAnimating As System.Threading.Thread
    Private _intAnimatingDelay As Integer = 0

    Public Sub New(frmToPass As Form, strZombieType As String, intSpotX As Integer, intSpotY As Integer)

        'Set
        _frmToPass = frmToPass

        'Set
        _strTypeOfZombie = strZombieType

        'Preset
        btmZombie = gbtmUnderwaterZombieGameLose

        'Set
        pntZombie = New Point(intSpotX, intSpotY)

    End Sub

    Public Sub Start(Optional intAnimatingDelay As Integer = 50, Optional blnSpawn As Boolean = True)

        'Set
        _intAnimatingDelay = intAnimatingDelay

        'Start thread
        thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
        thrAnimating.Start()

    End Sub

    Public ReadOnly Property ZombieImage() As Bitmap

        'Return
        Get
            Return btmZombie
        End Get

    End Property

    Public Property ZombiePoint() As Point

        'Return
        Get
            Return pntZombie
        End Get

        'Set
        Set(value As Point)
            pntZombie = value
        End Set

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

        'Continue
        While Not blnKeepUsingAnimatingThread

            'Sleep
            System.Threading.Thread.Sleep(_intAnimatingDelay)

            'Check frame
            Select Case _strTypeOfZombie
                Case "Underwater"
                    Select Case intFrame
                        Case 1 To 10
                            'Change y
                            ZombiePoint = New Point(ZombiePoint.X, ZombiePoint.Y + 25)
                            'Check
                            If intFrame = 10 Then
                                blnKeepUsingAnimatingThread = False
                            End If
                    End Select
            End Select

            Debug.Print("STILL ANIMATING")

            'Increase Frame
            intFrame += 1

        End While

    End Sub

End Class