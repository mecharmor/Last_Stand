'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class Character

    'Declare
    Private intFrame As Integer = 1
    Private thrAnimation As System.Threading.Thread
    Private intMovementPoint As Integer = 0
    Private blnShot As Boolean = False

    'Bitmaps
    Private btmCharacterStand1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/1.png"), 643, 710)
    Private btmCharacterStand2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/2.png"), 643, 710)
    Private btmCharacterShoot1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/3.png"), 643, 710)
    Private btmCharacterShoot2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/4.png"), 643, 710)
    Public btmCharacter As Bitmap
    Public rectCharacter As Rectangle

    Public Sub New(intSpawnPosition As Integer, Optional blnStart As Boolean = False)

        'Set
        intMovementPoint = intSpawnPosition
        rectCharacter = New Rectangle(intMovementPoint, 400, 643, 710)

        'Set
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

    Public Sub CharacterShot()

        'Change frame immediately
        btmCharacter = btmCharacterShoot1
        blnShot = True

    End Sub

    Private Sub Animating()

        'Loop
        While True
            'Frame
            Select Case intFrame
                Case 1
                    btmCharacter = btmCharacterStand1
                    'Reduce CPU usage
                    System.Threading.Thread.Sleep(300) '300
                Case 2
                    btmCharacter = btmCharacterStand2
                    'Reduce CPU usage
                    System.Threading.Thread.Sleep(300) '300
                Case 3
                    btmCharacter = btmCharacterShoot1
                    'Reduce CPU usage
                    System.Threading.Thread.Sleep(185) '300
                Case 4
                    btmCharacter = btmCharacterShoot2
                    'Reduce CPU usage
                    System.Threading.Thread.Sleep(160) '300
            End Select
            'Update the frame
            If blnShot Then
                'Change frame
                If intFrame = 3 Then
                    intFrame = 4
                    blnShot = False
                Else
                    intFrame = 3
                End If
            Else
                'Change frame
                If intFrame = 1 Then
                    intFrame = 2
                Else
                    intFrame = 1
                End If
            End If
        End While

    End Sub

    Public Sub StopCharacter()

        'Abort thread
        If thrAnimation.IsAlive Then
            thrAnimation.Abort()
        End If

    End Sub

End Class