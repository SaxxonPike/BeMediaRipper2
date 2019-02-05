Option Explicit On
Module modQueue
    Public SoundFileProgress As Integer
    Public DataFileProgress As Integer
    Public ThisJob As QueueJob
    Public Queue() As QueueJob
    Public QueueCount As Integer

    Public Structure QueueJob
        Public Name As String
        Public Enabled As Boolean
        Public GameConfig As Integer
        Public SourceFile As String
        Public SourceExtraFile As String
        Public TargetFolder As String
        Public RipKeysounds As Boolean
        Public RipBGM As Boolean
        Public RipCharts As Boolean
        Public RipVideos As Boolean
        Public RipGraphics As Boolean
        Public ConvertKeysounds As Boolean
        Public ConvertBGM As Boolean
        Public ConvertChart As Boolean
        Public ConvertGraphics As Boolean
        Public AutoName As Boolean
        Public AutoStructure As Boolean
        Public StripSilence As Boolean
        Public DontDecompress As Boolean
    End Structure

    Public Sub QueueGetNextJob()
        Dim x As Integer
        If QueueCount = 0 Then
            ThisJob.Enabled = False
            Exit Sub
        End If
        ThisJob = Queue(0)
        QueueCount -= 1
        For x = 0 To QueueCount - 1
            Queue(x) = Queue(x + 1)
        Next
        ReDim Preserve Queue(QueueCount - 1)
        ThisJob.Enabled = True
    End Sub

    Public Sub QueueAddJob(ByVal newJob As QueueJob)
        ReDim Preserve Queue(QueueCount)
        Queue(QueueCount) = newJob
        QueueCount += 1
    End Sub

    Public Sub QueueAddJobPriority(ByVal newJob As QueueJob)
        Dim x As Integer
        ReDim Preserve Queue(QueueCount)
        For x = QueueCount To 1 Step -1
            Queue(x) = Queue(x - 1)
        Next
        Queue(0) = newJob
        QueueCount += 1
    End Sub

    Public Sub QueueClear()
        QueueCount = 0
        ReDim Queue(0)
    End Sub

End Module
