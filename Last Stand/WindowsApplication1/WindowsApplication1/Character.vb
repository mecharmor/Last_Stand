'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class Character

    'Declare
    Private intFrame As Integer = 1
    Private thrWalking As System.Threading.Thread
    Private intMovementPoint As Integer = 0
    Public blnShot As Boolean = False

    'Bitmaps
    Private btmCharacterStand1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/1.png"), 643, 710)
    Private btmCharacterStand2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/2.png"), 643, 710)
    Private btmCharacterShoot1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/3.png"), 643, 710)
    Private btmCharacterShoot2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/4.png"), 643, 710)
    Public btmCharacter As Bitmap
    Public rectCharacter As Rectangle

    Public Sub New(intSpawnPosition As Integer)

        'Set
        intMovementPoint = intSpawnPosition
        rectCharacter = New Rectangle(intMovementPoint, 400, 643, 710)

        'Start thread
        thrWalking = New System.Threading.Thread(Sub() Walking())
        thrWalking.Start()

    End Sub

    Private Sub Walking()

        'Loop
        While True
            'Frame
            Select Case intFrame
                Case 1
                    btmCharacter = btmCharacterStand1
                Case 2
                    btmCharacter = btmCharacterStand2
                Case 3
                    btmCharacter = btmCharacterShoot1
                Case 4
                    btmCharacter = btmCharacterShoot2
            End Select
            'Update the frame
            If blnShot Then
                If intFrame = 3 Then
                    'Reduce CPU usage
                    System.Threading.Thread.Sleep(175) '250
                    'Change frame
                    intFrame = 4
                    blnShot = False
                Else
                    'Reduce CPU usage
                    System.Threading.Thread.Sleep(200) '300
                    'Change frame
                    intFrame = 3
                End If
            Else
                'Reduce CPU usage
                System.Threading.Thread.Sleep(225) '350
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
        If thrWalking.IsAlive Then
            thrWalking.Abort()
        End If

    End Sub

End Class