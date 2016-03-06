﻿'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class Zombie

    'Declare
    Private intFrame As Integer = 1
    Private thrWalking As System.Threading.Thread
    Private blnSwitch As Boolean = False
    Private intMovementPoint As Integer = 0

    'Bitmaps
    Private btmZombieWalk1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/1.png"), 484, 723)
    Private btmZombieWalk2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/2.png"), 484, 723)
    Private btmZombieWalk3 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/3.png"), 484, 723)
    Private btmZombieWalk4 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/4.png"), 484, 723)
    Public btmZombie As Bitmap
    Public rectZombie As Rectangle

    Public Sub New(intSpawnPosition As Integer)

        'Set
        intMovementPoint = intSpawnPosition

        'Start thread
        thrWalking = New System.Threading.Thread(Sub() Walking())
        thrWalking.Start()

    End Sub

    Private Sub Walking()

        'Loop
        While True
            'Change rectangle
            rectZombie = New Rectangle(intMovementPoint, 400, 484, 723)
            intMovementPoint -= 7
            'Frame
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
            'Update the frame
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
        End While

    End Sub

    Public Sub StopZombie()

        'Abort thread
        If thrWalking.IsAlive Then
            thrWalking.Abort()
        End If

    End Sub

End Class