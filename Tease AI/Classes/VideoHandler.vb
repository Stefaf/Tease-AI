
'#Const WMP = 1
#Const VLC = 1

Public Module VideoHandlerCreator
    Function CreateHandler() As VideoHandler
        If My.Application.CommandLineArgs.Contains("-VLC") Then
#If VLC Then
            Return New VLCVideoHandler
#Else
            Throw New ApplicationException("Trying to use VLC but this executable currently not support it!")
#End If
        ElseIf My.Application.CommandLineArgs.Contains("-WMP") Then
#If WMP Then
            Return New WMPVideoHandler
#Else
            Throw New ApplicationException("Trying to use WMP but this executable currently not support it!")
#End If
        Else
            Return New VideoHandler
        End If
    End Function

End Module


Public Class VideoHandler
    'Inherits System.Windows.Forms.Control
    Implements System.ComponentModel.ISupportInitialize

    Enum PlayState
        Undefined = 0   '<--- Used in VideoHandler
        Stopped = 1     '<--- Used in Form1
        Paused = 2      '<--- Used in Form1
        Playing = 3     '<--- Used in Form1
    End Enum

    Private FormControl As System.Windows.Forms.Control

    Public Event StateChanged(ByRef sender As Object, ByVal newState As PlayState)

    Protected Sub ThrowEventStateChanged(ByRef sender As Object, ByVal newState As PlayState)
        RaiseEvent StateChanged(sender, newState)
    End Sub

    Public Property Form As System.Windows.Forms.Control
        Get
            If FormControl Is Nothing Then
                FormControl = New System.Windows.Forms.Control
            End If
            Return FormControl
        End Get
        Protected Set(value As System.Windows.Forms.Control)
            FormControl = value
        End Set
    End Property

    Public Property Visible As Boolean
        Get
            Return Form.Visible
        End Get
        Set(value As Boolean)
            Form.Visible = value
        End Set
    End Property

    Public WriteOnly Property Name As String
        Set(value As String)
            Form.Name = value
        End Set
    End Property

    Public WriteOnly Property Enabled As Boolean
        Set(value As Boolean)
            Form.Enabled = value
        End Set
    End Property

    Public Overridable Property MediaURI As String
        Get
            Return ""
        End Get
        Set(value As String)
        End Set
    End Property

    Public Property Height As Integer
        Get
            Return Form.Height
        End Get
        Set(value As Integer)
            Form.Height = value
        End Set
    End Property

    Public Property Width As Integer
        Get
            Return Form.Width
        End Get
        Set(value As Integer)
            Form.Width = value
        End Set
    End Property

    Public WriteOnly Property Size As System.Drawing.Size
        Set(value As System.Drawing.Size)
            Form.Size = value
        End Set
    End Property

    Public Overridable Property ShowControls As Boolean
        Get
            Return False
        End Get
        Set(value As Boolean)
        End Set
    End Property

    Public WriteOnly Property Anchor As System.Windows.Forms.AnchorStyles
        Set(value As System.Windows.Forms.AnchorStyles)
            Form.Anchor = value
        End Set
    End Property

    Public WriteOnly Property Location As System.Drawing.Point
        Set(value As System.Drawing.Point)
            Form.Location = value
        End Set
    End Property

    Public Overridable ReadOnly Property CurrentMediaDuration As Double
        Get
            Return 0
        End Get
    End Property

    Public Overridable Property CurrentMediaCurrentPosition As Double
        Get
            Return 0
        End Get
        Set(value As Double)
        End Set
    End Property

    Public Overridable ReadOnly Property CurrentPlayListCount As Integer
        Get
            Return 0
        End Get
    End Property

    Public Overridable Sub ClearPlayList()
    End Sub

    Public Overridable WriteOnly Property Looped As Boolean
        Set(value As Boolean)
        End Set
    End Property

    Public Overridable Property Muted As Boolean
        Get
            Return True
        End Get
        Set(value As Boolean)
        End Set
    End Property

    Public Overridable Property Volume As Double
        Get
            Return 0
        End Get
        Set(value As Double)
        End Set
    End Property

    Public Overridable WriteOnly Property StretchToFit As Boolean
        Set(value As Boolean)
        End Set
    End Property

    Public WriteOnly Property TabIndex As Integer
        Set(value As Integer)
            Form.TabIndex = value
        End Set
    End Property

    Public Overridable Property State As PlayState
        Get
            Return PlayState.Undefined
        End Get
        Set(value As PlayState)
        End Set
    End Property

    Public Overridable Sub PlayIt()
    End Sub

    Public Overridable Sub PauseIt()
    End Sub

    Public Overridable Sub StopIt()   '[Note: Stop is a reserved key word]
    End Sub

    Public Overridable Sub BeginInit() Implements System.ComponentModel.ISupportInitialize.BeginInit
    End Sub
    Public Overridable Sub EndInit() Implements System.ComponentModel.ISupportInitialize.EndInit
    End Sub
End Class

#If WMP Then
Public Class WMPVideoHandler
    Inherits VideoHandler

    'Enum WMPPlayState
    '    wmppsUndefined = 0
    '    wmppsStopped = 1
    '    wmppsPaused = 2
    '    wmppsPlaying = 3
    '    wmppsScanForward = 4
    '    wmppsScanReverse = 5
    '    wmppsBuffering = 6
    '    wmppsWaiting = 7
    '    wmppsMediaEnded = 8
    '    wmppsTransitioning = 9
    '    wmppsReady = 10
    '    wmppsReconnecting = 11
    '    wmppsLast = 12
    'End Enum

    Dim WMPPlayer As AxWMPLib.AxWindowsMediaPlayer
    'Dim controlBar As Boolean

    Private Sub RaiseChangeState(Sender As Object, e As AxWMPLib._WMPOCXEvents_PlayStateChangeEvent) Handles WMPPlayer.PlayStateChange
        ThrowEventStateChanged(Sender, e.newState)
    End Sub

    Sub New()
        WMPPlayer = New AxWMPLib.AxWindowsMediaPlayer
        Form = WMPPlayer
        WMPPlayer.uiMode = "none"
        'controlBar = False
    End Sub

    Public Overrides Property ShowControls As Boolean
        Get
            'Return controlBar
            Return WMPPlayer.uiMode = "mini"
        End Get
        Set(value As Boolean)
            'controlBar = value
            If value Then
                'Height += 60
                WMPPlayer.uiMode = "mini"
            Else
                'Height -= 60
                WMPPlayer.uiMode = "none"
            End If
        End Set
    End Property

    Public Overrides Property MediaURI As String
        Get
            Return WMPPlayer.URL
        End Get
        Set(value As String)
            WMPPlayer.URL = value
        End Set
    End Property

    Public Overrides ReadOnly Property CurrentMediaDuration As Double
        Get
            Return WMPPlayer.currentMedia.duration
        End Get
    End Property

    Public Overrides Property CurrentMediaCurrentPosition As Double
        Get
            Return WMPPlayer.Ctlcontrols.currentPosition
        End Get
        Set(value As Double)
            WMPPlayer.Ctlcontrols.currentPosition = value
        End Set
    End Property

    Public Overrides ReadOnly Property CurrentPlayListCount As Integer
        Get
            Return WMPPlayer.currentPlaylist.count
        End Get
    End Property

    Public Overrides Sub ClearPlayList()
        WMPPlayer.currentPlaylist.Clear()
    End Sub

    Public Overrides WriteOnly Property Looped As Boolean
        Set(value As Boolean)
            WMPPlayer.settings.setMode("loop", value)
        End Set
    End Property

    Public Overrides Property Muted As Boolean
        Get
            Return WMPPlayer.settings.mute
        End Get
        Set(value As Boolean)
            WMPPlayer.settings.mute = value
        End Set
    End Property

    Public Overrides Property Volume As Double
        Get
            Return WMPPlayer.settings.volume
        End Get
        Set(value As Double)
            WMPPlayer.settings.volume = value
        End Set
    End Property

    Public Overrides WriteOnly Property StretchToFit As Boolean
        Set(value As Boolean)
            WMPPlayer.StretchToFit = value
        End Set
    End Property

    Public Overrides Property State As PlayState
        Get
            Return WMPPlayer.playState
        End Get
        Set(value As PlayState)
            '[JT-TODO: handle enum conversion here]
            WMPPlayer.playState = value
        End Set
    End Property

    Public Overrides Sub PlayIt()
        WMPPlayer.Ctlcontrols.play()
    End Sub

    Public Overrides Sub PauseIt()
        WMPPlayer.Ctlcontrols.pause()
    End Sub

    Public Overrides Sub StopIt()   '[Note: Stop is a reserved key word]
        WMPPlayer.Ctlcontrols.stop()
    End Sub

    Public Overrides Sub BeginInit()
        WMPPlayer.BeginInit()
    End Sub

    Public Overrides Sub EndInit()
        WMPPlayer.EndInit()
    End Sub
End Class
#End If

#If VLC Then
Public Class VLCVideoHandler
    Inherits VideoHandler

    Enum VLCPlayState
        VLCStopped = 5
        VLCPaused = 3
        VLCPlaying = 1
        'not all listed
    End Enum

    Dim WithEvents VLCPlayer As AxAXVLC.AxVLCPlugin2 'AXVLC.VLCPlugin2

    Dim LastPlayed As String

    Private Sub RaiseChangeStateToPlay() Handles VLCPlayer.MediaPlayerPlaying
        ThrowEventStateChanged(VLCPlayer, PlayState.Playing)
    End Sub

    Private Sub RaiseChangeStateToPause() Handles VLCPlayer.MediaPlayerPaused
        ThrowEventStateChanged(VLCPlayer, PlayState.Paused)
    End Sub

    Private Sub RaiseChangeStateToStop() Handles VLCPlayer.MediaPlayerStopped
        ThrowEventStateChanged(VLCPlayer, PlayState.Stopped)
    End Sub

    Sub New()
        VLCPlayer = New AxAXVLC.AxVLCPlugin2
        Form = VLCPlayer
        LastPlayed = ""
        'VLCPlayer.AutoPlay = True
    End Sub

    Public Overrides Property ShowControls As Boolean
        Get
            Return VLCPlayer.Toolbar
        End Get
        Set(value As Boolean)
            VLCPlayer.Toolbar = value
        End Set
    End Property

    Public Overrides Property MediaURI As String
        Get
            Return LastPlayed
        End Get
        Set(value As String)
            If VLCPlayer.playlist.items.count <> 0 Then
                If State = VideoHandler.PlayState.Playing Or State = VideoHandler.PlayState.Paused Then
                    StopIt()
                End If
                ClearPlayList()
            End If
            LastPlayed = value
            VLCPlayer.playlist.add("file:///" + LastPlayed)
            PlayIt()
        End Set
    End Property

    Public Overrides ReadOnly Property CurrentMediaDuration As Double
        Get
            Return VLCPlayer.input.length
        End Get
    End Property

    Public Overrides Property CurrentMediaCurrentPosition As Double
        Get
            Return VLCPlayer.input.position
        End Get
        Set(value As Double)
            VLCPlayer.input.position = value
        End Set
    End Property

    Public Overrides ReadOnly Property CurrentPlayListCount As Integer
        Get
            Return VLCPlayer.playlist.items.count
        End Get
    End Property

    Public Overrides Sub ClearPlayList()
        VLCPlayer.playlist.items.clear()
    End Sub

    Public Overrides WriteOnly Property Looped As Boolean
        Set(value As Boolean)
            VLCPlayer.AutoLoop = value
        End Set
    End Property

    Public Overrides Property Muted As Boolean
        Get
            Return VLCPlayer.audio.mute
        End Get
        Set(value As Boolean)
            VLCPlayer.audio.mute = value
        End Set
    End Property

    Public Overrides Property Volume As Double
        Get
            Return VLCPlayer.audio.volume
        End Get
        Set(value As Double)
            VLCPlayer.audio.volume = value
        End Set
    End Property

    Public Overrides WriteOnly Property StretchToFit As Boolean
        Set(value As Boolean)
            If True Then
                Size = Form.Parent.Size
            End If
        End Set
    End Property

    Public Overrides Property State As PlayState
        Get
            Select Case VLCPlayer.input.state
                Case 1 'VLCPlaying
                    Return PlayState.Playing
                Case 3 'VLCPaused
                    Return PlayState.Paused
                Case 5 'VLCStopped
                    Return PlayState.Stopped
            End Select
            Return VLCPlayer.input.state
        End Get
        Set(value As PlayState)
            Select Case value
                Case PlayState.Playing
                    VLCPlayer.playlist.play()
                Case PlayState.Paused
                    VLCPlayer.playlist.pause()
                Case PlayState.Stopped
                    VLCPlayer.playlist.stop()
            End Select
        End Set
    End Property

    Public Overrides Sub PlayIt()
        VLCPlayer.playlist.play()
    End Sub

    Public Overrides Sub PauseIt()
        VLCPlayer.playlist.pause()
    End Sub

    Public Overrides Sub StopIt()   '[Note: Stop is a reserved key word]
        VLCPlayer.playlist.stop()
    End Sub

    Public Overrides Sub BeginInit()
        VLCPlayer.BeginInit()
    End Sub

    Public Overrides Sub EndInit()
        VLCPlayer.EndInit()
    End Sub
End Class
#End If