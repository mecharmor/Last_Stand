﻿'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsZombie

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 1
    Private thrAnimation As System.Threading.Thread
    Private blnSwitch As Boolean = False
    Private intSpotX As Integer = 0
    Private intSpotY As Integer = 0
    Private _intSpeed As Integer = 25

    'Bitmaps
    Public btmZombie As Bitmap
    Public pntZombie As Point

    'Timer
    Private sttTimer As System.Timers.Timer

    'Death
    Private blnAlive As Boolean = True
    Private blnPaintOnBackgroundAfterDead As Boolean = False

    'Pinning character
    Private blnPinning As Boolean = False

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, intSpeed As Integer, Optional blnStart As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Preset
        btmZombie = gbtmZombieWalk(0)

        'Set
        _intSpeed = intSpeed

        'Set
        intSpotX = intSpawnX
        intSpotY = intSpawnY
        pntZombie = New Point(intSpotX, intSpotY)

        'Set frame timer
        sttTimer = New System.Timers.Timer(200) 'For walking
        AddHandler sttTimer.Elapsed, AddressOf Animating

        'Start
        If blnStart Then
            Start()
        End If

    End Sub

    Public Sub Start()

        'Enable timer
        sttTimer.Enabled = True

    End Sub

    Private Sub Animating(sender As Object, e As System.Timers.ElapsedEventArgs)

        'Check if alive before checking frame
        If blnAlive Then
            'Check if pinning or walking
            If blnPinning Then
                'Pinning the character
                Select Case intFrame 'Interval is 200 in this area
                    Case 1
                        btmZombie = gbtmZombiePin(0)
                        intFrame = 2
                    Case 2
                        btmZombie = gbtmZombiePin(1)
                        intFrame = 1
                End Select
            Else
                'Change point
                pntZombie.X = intSpotX
                intSpotX -= _intSpeed 'Speed they come at
                'Check if going forward or backwards with the frames (1, 2, 3, 4) or (3, 2, 1) but becomes (1, 2, 3, 4, 3, 2, 1, 2, 3, 4, etc.)
                If Not blnSwitch Then
                    'Check frame
                    Select Case intFrame 'Interval is 300 in this area
                        Case 1
                            btmZombie = gbtmZombieWalk(0)
                            intFrame = 2
                        Case 2
                            btmZombie = gbtmZombieWalk(1)
                            intFrame = 3
                        Case 3
                            btmZombie = gbtmZombieWalk(2)
                            intFrame = 4
                        Case 4
                            btmZombie = gbtmZombieWalk(3)
                            intFrame = 1
                            'Switch up time
                            blnSwitch = True
                    End Select
                Else
                    'Check frame
                    Select Case intFrame
                        Case 1
                            btmZombie = gbtmZombieWalk(2)
                            intFrame = 2
                        Case 2
                            btmZombie = gbtmZombieWalk(1)
                            intFrame = 1
                            'Switch up time
                            blnSwitch = False
                    End Select
                End If
            End If
        Else 'Interval is 150 in this area
            'Check frame
            Select Case intFrame
                Case 1
                    btmZombie = gbtmZombieDeath(1)
                    intFrame = 2
                Case 2
                    btmZombie = gbtmZombieDeath(2)
                    intFrame = 3
                Case 3
                    btmZombie = gbtmZombieDeath(3)
                    intFrame = 4
                Case 4
                    btmZombie = gbtmZombieDeath(4)
                    intFrame = 5
                Case 5
                    btmZombie = gbtmZombieDeath(5)
                    'Stop timer and handler
                    StopAndDispose()
                    'Paint on top of the background
                    blnPaintOnBackgroundAfterDead = True
            End Select

        End If

    End Sub

    Public ReadOnly Property PaintOnBackgroundAfterDead() As Boolean

        'Return if ready to paint after death
        Get
            Return blnPaintOnBackgroundAfterDead
        End Get

    End Property

    Public ReadOnly Property IsPinning() As Boolean

        'Return pinning or not
        Get
            Return blnPinning
        End Get

    End Property

    Public ReadOnly Property IsAlive() As Boolean

        'Return alive or not
        Get
            Return blnAlive
        End Get

    End Property

    Public Sub Dead()

        'Stop timer and change interval
        sttTimer.Enabled = False
        sttTimer.Interval = 150

        'Set dead
        blnAlive = False

        'Change frame immediately
        intFrame = 1
        btmZombie = gbtmZombieDeath(0)

        'Declare
        Dim rndNumber As New Random

        'Play sound of death
        Dim udcDeathSound As New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "/Sounds/ZombieDeath" & CStr(rndNumber.Next(1, 4)) &
                                          ".mp3", 2000, gintSoundVolume)

        'Start timer again
        sttTimer.Enabled = True

    End Sub

    Public Sub Pin()

        'Stop timer and change interval
        sttTimer.Enabled = False
        sttTimer.Interval = 200

        'Set pinning
        blnPinning = True

        'Change frame immediately
        intFrame = 1
        btmZombie = gbtmZombiePin(0)

        'Start timer again
        sttTimer.Enabled = True

    End Sub

    Public Sub StopAndDispose()

        'Stop timer
        sttTimer.Enabled = False

        'Remove handler
        RemoveHandler sttTimer.Elapsed, AddressOf Animating

        'Dispose
        sttTimer.Dispose()

    End Sub

End Class