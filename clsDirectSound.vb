Imports Microsoft.DirectX.DirectSound
Imports Microsoft.DirectX.AudioVideoPlayback


Public Class clsDirectSound
    Private Const HighestSoundIndex As Integer = 4095
    Private Enum SoundAPIType
        None
        DirectSound
        AudioVideoPlayback
    End Enum

    Private Initted As Boolean
    Private _dev As Device
    Private Sounds(HighestSoundIndex) As Object
    Private SoundLoaded(HighestSoundIndex) As Boolean
    Private SoundAPI(HighestSoundIndex) As SoundAPIType

    Private ThisDXSound As SecondaryBuffer
    Private ThisAVSound As Audio


    Public Function Init(ByVal hwnd As System.IntPtr) As Boolean
        Try
            _dev = New Device
            _dev.SetCooperativeLevel(hwnd, CooperativeLevel.Normal)
            Initted = True
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function IsReady() As Boolean
        Return Initted
    End Function

    Public Function SoundLoad(ByVal sFileName As String, ByVal iNumber As Integer) As Boolean
        Dim BufferDesc As New BufferDescription
        BufferDesc.ControlEffects = False
        If My.Computer.FileSystem.FileExists(sFileName) = False Then
            Return False
        End If
        If iNumber >= 0 And iNumber <= HighestSoundIndex Then
            SoundLoaded(iNumber) = False
            SoundAPI(iNumber) = SoundAPIType.None
            'first load with DirectX since it is better at playback
            Try
                Sounds(iNumber) = New SecondaryBuffer(sFileName, BufferDesc, _dev)
                SoundLoaded(iNumber) = True
                SoundAPI(iNumber) = SoundAPIType.DirectSound
                Return True
            Catch ex As Exception
                'didn't load with DirectX, could be an unsupported format so let's use AVPlayback
                Try
                    Sounds(iNumber) = New Audio(sFileName)
                    SoundLoaded(iNumber) = True
                    SoundAPI(iNumber) = SoundAPIType.AudioVideoPlayback
                    Return True
                Catch ex2 As Exception
                    'didn't load with anything
                    SoundLoaded(iNumber) = False
                End Try
                Return SoundLoaded(iNumber)
            End Try
        End If
        Return False
    End Function

    Public Function SoundPlay(ByVal iNumber As Integer) As Boolean
        If (iNumber >= 0 And iNumber <= HighestSoundIndex) Then
            If SoundLoaded(iNumber) Then
                Try
                    Select Case SoundAPI(iNumber)
                        Case SoundAPIType.DirectSound
                            ThisDXSound = Sounds(iNumber)
                            If ThisDXSound.Status.Playing Then
                                ThisDXSound.Stop()
                                ThisDXSound.SetCurrentPosition(0)
                                ThisDXSound.Play(0, BufferPlayFlags.Default)
                            Else
                                ThisDXSound.Play(0, BufferPlayFlags.Default)
                            End If
                        Case SoundAPIType.AudioVideoPlayback
                            ThisAVSound = Sounds(iNumber)
                            If ThisAVSound.Playing Then
                                ThisAVSound.Stop()
                            End If
                            ThisAVSound.Play()
                    End Select
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            Else
                Return False
            End If
        End If
        Return False
    End Function

    Public Sub SoundStop(ByVal iNumber As Integer)
        If (iNumber >= 0 And iNumber <= HighestSoundIndex) Then
            Select Case SoundAPI(iNumber)
                Case SoundAPIType.DirectSound
                    ThisDXSound = Sounds(iNumber)
                    ThisDXSound.Stop()
                    ThisDXSound.SetCurrentPosition(0)
                Case SoundAPIType.AudioVideoPlayback
                    ThisAVSound = Sounds(iNumber)
                    ThisAVSound.Stop()
            End Select
        End If
    End Sub

    Public Sub SoundFree(ByVal iNumber As Integer)
        If iNumber >= 0 And iNumber <= HighestSoundIndex Then
            SoundStop(iNumber)
            Sounds(iNumber) = Nothing
            SoundLoaded(iNumber) = False
            SoundAPI(iNumber) = SoundAPIType.None
        End If
    End Sub

    Public Sub SoundFreeAll()
        Dim x As Integer
        For x = 0 To HighestSoundIndex
            If SoundLoaded(x) Then
                SoundFree(x)
            End If
        Next
    End Sub

End Class
