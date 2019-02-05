Option Explicit On
Module modChart

    Public Const ChartMaxSize = 65536 'if no filesize is given, use this instead
    Public ChartForceTiming As Integer = 0

    'universal chart format
    Public Structure ChartFormat
        Public OffsetBase As Long
        Public OffsetMetric As Double
        Public OffsetWithinMeasure As Double
        Public OffsetMSec As Long
        Public OffsetBaseWithinMeasure As Integer
        Public Measure As Integer
        Public Value As Double
        Public Lane As Integer
        Public Parameter As ChartParameter
        Public Player As Integer
        Public BMSUsed As Boolean
    End Structure

    Public Enum ChartParameter As Integer
        None
        Measure
        TempoChange
        SoundChange
        Note
        FreezeNote
        BGM
        Meter
        Info
        Overlay
        EndSong
        Judgement
        Bad
    End Enum

    Public Structure ChartStruct
        Public EventCount As Long
        Public OffsetRate As Double
        Public Events() As ChartFormat
        Public SongLengthSeconds As Integer
        Public MeasureMetrics() As Double
        Public MeasurePreferredUnits() As Integer
        Public MeasureOffsets() As Integer
        Public WaveTable() As String
        Public ChartType As ChartTypes
        Public HasProgramChanges As Boolean
        Public Sub Init()
            EventCount = 0
            OffsetRate = 0
            ReDim Events(0)
            SongLengthSeconds = 0
            ReDim MeasureMetrics(0)
            ReDim MeasurePreferredUnits(0)
            ReDim MeasureOffsets(0)
            ReDim WaveTable(0)
            ChartType = ChartTypes.Undefined
            HasProgramChanges = False
        End Sub
        Public Function GetUsedSamples() As Integer()
            Dim x As Integer
            Dim b(4096) As Boolean
            Dim i As Integer = 0
            Dim r() As Integer = {}
            For x = 0 To UBound(Events)
                With Events(x)
                    If (.Parameter = ChartParameter.BGM) Or (.Parameter = ChartParameter.Note) Then
                        If b(.Value) = False Then
                            b(.Value) = True
                            If i = 0 Then
                                ReDim r(0)
                            Else
                                ReDim Preserve r(i)
                            End If
                            r(i) = .Value
                            i += 1
                        End If
                    End If
                End With
            Next
            Return r
        End Function
        Public Function NoteCount() As Integer
            NoteCount = 0
            Dim x As Integer
            For x = 0 To UBound(Events)
                With Events(x)
                    If .Parameter = ChartParameter.Note Then
                        NoteCount += 1
                    End If
                End With
            Next
        End Function
        Public Function NoteCount(ByVal iPlayer As Integer) As Integer
            NoteCount = 0
            Dim x As Integer
            For x = 0 To UBound(Events)
                With Events(x)
                    If .Parameter = ChartParameter.Note Then
                        If .Player = iPlayer Then
                            NoteCount += 1
                        End If
                    End If
                End With
            Next
        End Function
        Public Sub AdjustAllNotes(ByVal AdjustAmount As Integer)
            For x = 0 To UBound(Events)
                With Events(x)
                    If .Parameter = ChartParameter.BGM Or .Parameter = ChartParameter.Note Or .Parameter = ChartParameter.SoundChange Then
                        If .Value > 0 Then
                            .Value += AdjustAmount
                        End If
                    End If
                End With
            Next
        End Sub
        Public Sub ConvertAllNotes(ByVal ConvertFrom As Integer, ByVal ConvertTo As Integer)
            For x = 0 To UBound(Events)
                With Events(x)
                    If .Parameter = ChartParameter.BGM Or .Parameter = ChartParameter.Note Or .Parameter = ChartParameter.SoundChange Then
                        If .Value = ConvertFrom Then
                            .Value = ConvertTo
                        End If
                    End If
                End With
            Next
        End Sub
        Public Function SelectNotes(ByRef ReturnNotes() As ChartFormat, Optional ByVal iParameter As ChartParameter = -1, Optional ByVal iMeasure As Integer = -1, Optional ByVal iLane As Integer = -1, Optional ByVal iPlayer As Integer = -1) As Boolean
            Dim ReturnNoteCount As Integer = 0
            Dim e As Boolean
            Dim x As Integer
            ReDim ReturnNotes(0)
            ReturnNotes(0).Parameter = ChartParameter.None
            For x = 0 To UBound(Events)
                With Events(x)
                    e = True
                    If iMeasure > -1 Then
                        If .Measure <> iMeasure Then
                            e = False
                        End If
                    End If
                    If iLane > -1 Then
                        If .Lane <> iLane Then
                            e = False
                        End If
                    End If
                    If iParameter > -1 Then
                        If .Parameter <> iParameter Then
                            e = False
                        End If
                    End If
                    If iPlayer > -1 Then
                        If .Player <> iPlayer Then
                            e = False
                        End If
                    End If
                    If e Then
                        ReDim Preserve ReturnNotes(ReturnNoteCount)
                        ReturnNotes(ReturnNoteCount) = Events(x)
                        ReturnNoteCount += 1
                    End If
                End With
            Next
            Return (ReturnNoteCount > 0)
        End Function
    End Structure

    Public Enum ChartTypes As Integer
        Undefined
        IIDXAC
        IIDXCS
        IIDXCS2
        IIDXCS5
        Popn4
        Popn6
        Popn8
        BME
        PMS
    End Enum

    Public Chart As ChartStruct
    Public ChartRaw() As Byte

    'load a chart from a file
    Public Sub ChartLoadFile(ByVal sFileName As String, ByVal iChartType As ChartTypes, Optional ByVal lOffset As Long = 0, Optional ByVal lLength As Long = -1, Optional ByVal iSubChart As Integer = 0)
        Dim bChartBytes() As Byte
        If lLength < 0 Then
            If FileExists(sFileName) Then
                lLength = My.Computer.FileSystem.GetFileInfo(sFileName).Length
            End If
        End If
        ReDim bChartBytes(lLength - 1)
        If FileExists(sFileName) Then
            FileLoadMemory(bChartBytes, sFileName, lOffset, lLength)
            ChartLoad(bChartBytes, iChartType, , iSubChart)
        End If
    End Sub

    'load a chart from memory
    Public Sub ChartLoadMemory(ByRef bChartBytes() As Byte, ByVal iChartType As ChartTypes)
        ChartLoad(bChartBytes, iChartType)
    End Sub

    Public Sub ChartLoadFromOpen(ByVal FileNumber As Integer, ByVal Offset As Long, ByVal Encoding As DataEncodingType, ByVal ChartType As ChartTypes, Optional ByVal Length As Integer = -1, Optional ByVal DontConvert As Boolean = False)
        Dim cb() As Byte = {0}
        Dim db() As Byte = {0}
        If Length < 0 Then
            If (LOF(FileNumber) - Offset) > ChartMaxSize Then
                ReDim cb(ChartMaxSize - 1)
            Else
                ReDim cb(LOF(FileNumber) - Offset)
            End If
        Else
            ReDim cb(Length - 1)
        End If
        FileSystem.FileGet(FileNumber, cb, Offset + 1)
        Select Case Encoding
            Case DataEncodingType.None
                If (Not DontConvert) Then
                    ChartLoadMemory(cb, ChartType)
                End If
                ChartRaw = cb
            Case DataEncodingType.KonamiLZ77
                DataDecodeBemaniLZ77(cb, db)
                If (Not DontConvert) Then
                    ChartLoadMemory(db, ChartType)
                End If
                ChartRaw = db
        End Select
    End Sub

    'main load routine
    'returns true if there was a problem

    'todo: add popn support. timing for old popn is (50000 / 3)

    Private Function ChartReadNextLine(ByRef Data() As Byte, ByRef Offset As Integer) As String
        ChartReadNextLine = ""
        Do While Offset <= UBound(Data)
            If Data(Offset) >= &H20 Then
                ChartReadNextLine &= Chr(Data(Offset))
            ElseIf Data(Offset) = &HD Then
                Offset += 1
                Exit Do
            End If
            Offset += 1
        Loop
    End Function


    Private Function ChartLoad(ByRef ChartData() As Byte, ByVal iChartType As ChartTypes, Optional ByVal iForceTiming As Integer = -1, Optional ByVal iSubChart As Integer = 0) As Boolean
        ReDim Chart.WaveTable(0 To 1295)
        Dim ChartOffset As Integer
        Dim ValidNote As Boolean = False
        Dim TickRate As Integer = 0
        Dim ThisBPM As Double = 0
        Dim LastReadBPM As Integer = 0
        Dim BGMused As Boolean = False
        Dim ThisMeasure As Integer = 0
        Dim ThisMetric As Double = 0
        Dim ThisMeasureDefaultSize As Double = 0
        Dim ReferenceBase As Integer = 0
        Dim MeasureMetrics() As Double = {1}
        Dim MeasureMetricOffsets() As Double = {0}
        Dim MeasureCount As Integer = 0
        Dim LastMeasureOffset As Integer = 0
        Dim Val1 As Integer = 0
        Dim Val2 As Integer = 0
        Dim Val3 As Integer = 0
        Dim BPMTable(1295) As Double
        Dim ThisLane As Integer = 0
        Dim ThisPlayer As Integer = 0
        Dim ThisLine As String
        Dim ThisCommand As String
        Dim ThisParameter As String
        Dim ActiveKeysounds(255) As Integer
        Dim LastNoteOffset(255) As Integer
        Dim x As Integer
        Dim y As Integer
        Dim e As Boolean
        Dim HasWaveTable As Boolean = False
        Dim TempEvent As ChartFormat
        Dim ChartSize As Integer = UBound(ChartData) + 1
        Dim HasBPMChange As Boolean
        Dim HasNotes As Boolean
        Dim ConvertToMetric As Boolean
        ChartLoad = True
        If ChartSize < 4 Then
            Exit Function
        End If
        Chart.Init()
        SoundFileProgress = 0

        'LOAD STAGE ***********************************************************

        If iChartType = ChartTypes.IIDXCS Then
            If (ChartData(0) <> 8) And (ChartData(ChartSize - 4) = &HFF) And (ChartData(ChartSize - 3) = &H7F) And (ChartData(ChartSize - 2) = 0) And (ChartData(ChartSize - 1) = 0) Then
                iChartType = ChartTypes.IIDXCS2
            End If
        End If

        Select Case iChartType
            Case ChartTypes.BME, ChartTypes.PMS
                HasWaveTable = True
                ReDim MeasureMetrics(999)
                ReDim Chart.WaveTable(1295)
                ConvertToMetric = False
                Chart.EventCount = 0
                Chart.HasProgramChanges = False
                ThisMetric = 0
                ChartOffset = 0
                For x = 0 To 999
                    MeasureMetrics(x) = 1
                Next
                Do While ChartOffset < ChartSize
                    ThisLine = ChartReadNextLine(ChartData, ChartOffset)
                    If Left(ThisLine, 1) = "#" Then
                        If InStr(ThisLine, " ") > 0 Then
                            ThisCommand = UCase(Mid(ThisLine, 2))
                            ThisCommand = Left(ThisCommand, InStr(ThisCommand, " ") - 1)
                            ThisParameter = Mid(ThisLine, 3 + Len(ThisCommand))
                            If Left(ThisCommand, 3) = "BPM" Then
                                If ThisCommand = "BPM" Then
                                    ThisBPM = Val(ThisParameter)
                                Else
                                    BPMTable(DataUnBMEString(Mid(ThisCommand, 4, 2))) = Val(ThisParameter)
                                End If
                            ElseIf Left(ThisCommand, 3) = "WAV" Then
                                Chart.WaveTable(DataUnBMEString(Mid(ThisCommand, 4, 2))) = ThisParameter
                            End If
                        ElseIf Mid(ThisLine, 7, 1) = ":" Then
                            HasNotes = False
                            ThisCommand = UCase(Mid(ThisLine, 2))
                            ThisCommand = Left(ThisCommand, InStr(ThisCommand, ":") - 1)
                            ThisParameter = Mid(ThisLine, 3 + Len(ThisCommand))
                            ThisMeasure = Val(Left(ThisCommand, 3))
                            If ThisMeasure > MeasureCount Then
                                MeasureCount = ThisMeasure
                            End If
                            Select Case Mid(ThisCommand, 4, 2)
                                Case "02"
                                    'meter
                                    MeasureMetrics(ThisMeasure) = Val(ThisParameter)
                                Case "03"
                                    'old 01-FF bpm
                                    HasNotes = True
                                    HasBPMChange = False
                                    With TempEvent
                                        .Lane = 1
                                        .Parameter = ChartParameter.TempoChange
                                    End With
                                Case "08"
                                    'new bpm from list
                                    HasNotes = True
                                    HasBPMChange = True
                                    With TempEvent
                                        .Lane = 1
                                        .Parameter = ChartParameter.TempoChange
                                    End With
                                Case "01"
                                    'bgm
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 0
                                        .Parameter = ChartParameter.BGM
                                    End With
                                Case "11"
                                    'p1k1
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 0
                                        .Player = 1
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "12"
                                    'p1k2
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 1
                                        .Player = 1
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "13"
                                    'p1k3
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 2
                                        .Player = 1
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "14"
                                    'p1k4
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 3
                                        .Player = 1
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "15"
                                    'p1k5
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 4
                                        .Player = 1
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "18"
                                    'p1k6
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 5
                                        .Player = 1
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "19"
                                    'p1k7
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 6
                                        .Player = 1
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "16"
                                    'p1S
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 7
                                        .Player = 1
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "17"
                                    'p1FZ
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 8
                                        .Player = 1
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "21"
                                    'p2k1
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 0
                                        .Player = 2
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "22"
                                    'p2k2
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 1
                                        .Player = 2
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "23"
                                    'p2k3
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 2
                                        .Player = 2
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "24"
                                    'p2k4
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 3
                                        .Player = 2
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "25"
                                    'p2k5
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 4
                                        .Player = 2
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "28"
                                    'p2k6
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 5
                                        .Player = 2
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "29"
                                    'p2k7
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 6
                                        .Player = 2
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "26"
                                    'p2S
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 7
                                        .Player = 2
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "27"
                                    'p2FZ
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 8
                                        .Player = 2
                                        .Parameter = ChartParameter.Note
                                    End With
                                Case "51"
                                    'p2f1
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 0
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "52"
                                    'p1f2
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 1
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "53"
                                    'p1f3
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 2
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "54"
                                    'p1f4
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 3
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "55"
                                    'p1f5
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 4
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "58"
                                    'p1f6
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 5
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "59"
                                    'p1f7
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 6
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "56"
                                    'p1fS
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 7
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "61"
                                    'p2f1
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 0
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "62"
                                    'p2f2
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 1
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "63"
                                    'p2f3
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 2
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "64"
                                    'p2f4
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 3
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "65"
                                    'p2f5
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 4
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "68"
                                    'p2f6
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 5
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "69"
                                    'p2f7
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 6
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                                Case "66"
                                    'p2fS
                                    HasNotes = True
                                    With TempEvent
                                        .Lane = 7
                                        .Player = 2
                                        .Parameter = ChartParameter.FreezeNote
                                    End With
                            End Select
                            If HasNotes Then
                                With TempEvent
                                    For x = 1 To Len(ThisParameter) Step 2
                                        If Mid(ThisParameter, x, 2) <> "00" Then
                                            .OffsetWithinMeasure = ((x - 1) \ 2) / (Len(ThisParameter) \ 2)
                                            .Measure = ThisMeasure
                                            If .Parameter = ChartParameter.TempoChange Then
                                                If HasBPMChange = False Then
                                                    .Value = DataUnBMEString(Mid(ThisParameter, x, 2), "0123456789ABCDEF")
                                                Else
                                                    .Value = BPMTable(DataUnBMEString(Mid(ThisParameter, x, 2)))
                                                End If
                                            Else
                                                .Value = DataUnBMEString(Mid(ThisParameter, x, 2))
                                            End If
                                            ReDim Preserve Chart.Events(Chart.EventCount)
                                            Chart.Events(Chart.EventCount) = TempEvent
                                            Chart.EventCount += 1
                                        End If
                                    Next
                                End With
                            End If
                        End If
                    End If
                Loop
                'MeasureCount += 1
                ReDim Preserve MeasureMetrics(MeasureCount)
                ReDim MeasureMetricOffsets(MeasureCount)
                ReDim Chart.MeasurePreferredUnits(MeasureCount)

                For x = 0 To UBound(Chart.Events)
                    With Chart.Events(x)
                        If .Parameter = ChartParameter.Note Then

                        End If
                    End With
                Next

                'add song end event
                ReDim Preserve Chart.Events(Chart.EventCount)
                With TempEvent
                    .Lane = 0
                    .Measure = MeasureCount
                    .Parameter = ChartParameter.EndSong
                    .Player = 0
                    .OffsetWithinMeasure = 0
                    .Value = 0
                End With
                Chart.Events(Chart.EventCount) = TempEvent
                Chart.EventCount += 1

                'add BPM at beginning
                ReDim Preserve Chart.Events(Chart.EventCount)
                With TempEvent
                    .Lane = 1
                    .Measure = 0
                    .Parameter = ChartParameter.TempoChange
                    .Player = 0
                    .OffsetWithinMeasure = 0
                    .Value = ThisBPM
                End With
                Chart.Events(Chart.EventCount) = TempEvent
                Chart.EventCount += 1

                'add measures for the whole song
                For x = 1 To MeasureCount - 1
                    ReDim Preserve Chart.Events(Chart.EventCount)
                    With TempEvent
                        .Lane = 0
                        .Measure = x
                        .Parameter = ChartParameter.Measure
                        .Player = 0
                        .OffsetWithinMeasure = 0
                        .Value = 0
                    End With
                    Chart.Events(Chart.EventCount) = TempEvent
                    Chart.EventCount += 1
                Next

                'fill in missing metrics info now
                For x = 1 To UBound(MeasureMetricOffsets)
                    MeasureMetricOffsets(x) = MeasureMetricOffsets(x - 1) + MeasureMetrics(x - 1)
                Next
                For x = 0 To UBound(Chart.Events)
                    With Chart.Events(x)
                        .OffsetMetric = (.OffsetWithinMeasure * MeasureMetrics(.Measure)) + MeasureMetricOffsets(.Measure)
                    End With
                Next


            Case ChartTypes.IIDXCS5 '==========================================
                Chart.HasProgramChanges = True
                ConvertToMetric = True
                If Not (ChartData(ChartSize - 4) = &HFF And ChartData(ChartSize - 3) = &H7F And ChartData(ChartSize - 2) = 0 And ChartData(ChartSize - 1) = 0) Then
                    DebugLog("ChartLoad", "Chart signature is invalid.", "IIDXCS5")
                    Return True
                    Exit Function
                End If
                TickRate = (1000000 \ 58)
                If ChartData(ChartOffset + 2) = 0 Then
                    Do
                        If (ChartData(ChartOffset + 3) < 250) Then
                            ChartOffset += 4
                            Exit Do
                        End If
                        ChartOffset += 4
                    Loop While ChartData(ChartOffset + 2) = 0
                End If
                If ChartData(ChartOffset + 2) = &H10 Then
                    Do
                        If (ChartData(ChartOffset + 3) < 250) Then
                            ChartOffset += 4
                            Exit Do
                        End If
                        ChartOffset += 4
                    Loop While ChartData(ChartOffset + 2) = &H10
                End If
                With Chart
                    Do While ChartOffset < UBound(ChartData)
                        ValidNote = True
                        ReDim Preserve .Events(.EventCount)
                        With .Events(.EventCount)
                            .OffsetBase = DataMakeInt32(ChartData(ChartOffset), ChartData(ChartOffset + 1))
                            .BMSUsed = False
                            Val1 = ChartData(ChartOffset + 2) And &HF
                            Val2 = ChartData(ChartOffset + 2) >> 4
                            Val3 = ChartData(ChartOffset + 3)
                            Select Case Val1
                                Case 0
                                    Select Case Val2
                                        Case 0, 2, 4, 6, 8 '5keys
                                            .Player = 1
                                            .Lane = Val2 \ 2
                                        Case 1, 3, 5, 7, 9
                                            .Player = 2
                                            .Lane = (Val2 - 1) \ 2
                                        Case 10, 11 'scratch
                                            .Player = Val2 - 9
                                            .Lane = 7
                                        Case 12 'measure
                                            .Player = 1
                                            .Lane = 0
                                        Case 14, 15 'freescratch
                                            .Player = Val2 - 13
                                            .Lane = 8
                                        Case Else
                                            ValidNote = False
                                    End Select
                                    If Val2 <> 12 And Val2 <> 13 Then
                                        .Parameter = ChartParameter.Note
                                        .Value = ActiveKeysounds(.Lane + (128 * (.Player - 1)))
                                        If .Value = 0 Then
                                            ValidNote = False
                                        End If
                                    Else
                                        .Parameter = ChartParameter.Measure
                                    End If
                                Case 1
                                    Select Case Val2
                                        Case 0, 2, 4, 6, 8
                                            .Player = 1
                                            .Lane = Val2 \ 2
                                        Case 1, 3, 5, 7, 9
                                            .Player = 2
                                            .Lane = (Val2 - 1) \ 2
                                        Case 10, 11
                                            .Player = Val2 - 9
                                            .Lane = 7
                                        Case 14, 15
                                            .Player = Val2 - 13
                                            .Lane = 7
                                        Case Else
                                            ValidNote = False
                                    End Select
                                    .Parameter = ChartParameter.SoundChange
                                    If ValidNote Then
                                        .Value = Val3
                                        ActiveKeysounds(.Lane + (128 * (.Player - 1))) = Val3
                                        If .Lane = 7 Then
                                            ActiveKeysounds(8 + (128 * (.Player - 1))) = Val3
                                        End If
                                    End If
                                Case 2
                                    .Parameter = ChartParameter.TempoChange
                                    .Value = Val3 + (Val2 * 256)
                                    If ThisBPM = 0 Then
                                        ThisBPM = .Value
                                    End If
                                    .Lane = 1
                                    If LastReadBPM = .Value Then
                                        ValidNote = False
                                    End If
                                    LastReadBPM = .Value
                                Case 3
                                    'unknown...
                                Case 4
                                    If Val2 = 0 Then
                                        .Parameter = ChartParameter.EndSong
                                        Chart.SongLengthSeconds = (.OffsetBase * TickRate) \ 1000
                                    End If
                                Case 5
                                    .Parameter = ChartParameter.BGM
                                    .Value = Val3
                                Case 6
                                    .Parameter = ChartParameter.Judgement
                                    .Value = Val3
                                    .Lane = Val2
                                Case Else
                                    ValidNote = False
                            End Select
                            If .OffsetBase = &H7FFF Then
                                ValidNote = False
                                .Parameter = ChartParameter.Bad
                            End If
                        End With
                        ChartOffset += 4
                        If ValidNote Then
                            .EventCount += 1
                        End If
                    Loop
                End With


            Case ChartTypes.IIDXCS2 '==========================================
                Chart.HasProgramChanges = True
                ConvertToMetric = True
                If ChartData(0) <> 8 Then
                    If Not (ChartData(ChartSize - 4) = &HFF And ChartData(ChartSize - 3) = &H7F And ChartData(ChartSize - 2) = 0 And ChartData(ChartSize - 1) = 0) Then
                        DebugLog("ChartLoad", "Chart signature is invalid.", "IIDXCS2")
                        Return True
                        Exit Function
                    End If
                    ChartOffset = 0
                    TickRate = 16718
                Else
                    ChartOffset = 8
                    TickRate = DataMakeInt32(ChartData(4), ChartData(5), ChartData(6), ChartData(7))
                End If
                'bypass note counts...
                If ChartData(ChartOffset + 2) = 0 Then
                    Do
                        If (ChartData(ChartOffset + 3) < 250) Then
                            ChartOffset += 4
                            Exit Do
                        End If
                        ChartOffset += 4
                    Loop While ChartData(ChartOffset + 2) = 0
                End If
                If ChartData(ChartOffset + 2) = 1 Then
                    Do
                        If (ChartData(ChartOffset + 3) < 250) Then
                            ChartOffset += 4
                            Exit Do
                        End If
                        ChartOffset += 4
                    Loop While ChartData(ChartOffset + 2) = 1
                End If
                With Chart
                    Do While ChartOffset < UBound(ChartData)
                        ValidNote = True
                        ReDim Preserve .Events(.EventCount)
                        With .Events(.EventCount)
                            .OffsetBase = DataMakeInt32(ChartData(ChartOffset), ChartData(ChartOffset + 1))
                            .BMSUsed = False
                            Val1 = ChartData(ChartOffset + 2) And &HF
                            Val2 = ChartData(ChartOffset + 2) >> 4
                            Val3 = ChartData(ChartOffset + 3)
                            Select Case Val1
                                Case 0
                                    .Player = 1
                                    .Parameter = ChartParameter.Note
                                    .Lane = Val2
                                    If .OffsetBase = 0 Then
                                        ValidNote = False
                                    End If
                                    .Value = ActiveKeysounds(.Lane)
                                    If .Value = 0 Then
                                        ValidNote = False
                                    End If
                                Case 1
                                    .Player = 2
                                    .Parameter = ChartParameter.Note
                                    .Lane = Val2
                                    If .OffsetBase = 0 Then
                                        ValidNote = False
                                    End If
                                    .Value = ActiveKeysounds(.Lane + 128)
                                    If .Value = 0 Then
                                        ValidNote = False
                                    End If
                                Case 2
                                    .Player = 1
                                    .Parameter = ChartParameter.SoundChange
                                    .Value = Val3
                                    .Lane = Val2
                                    ActiveKeysounds(.Lane) = .Value
                                    If .Lane = 7 Then
                                        ActiveKeysounds(8) = .Value
                                    End If
                                Case 3
                                    .Player = 2
                                    .Parameter = ChartParameter.SoundChange
                                    .Value = Val3
                                    .Lane = Val2
                                    ActiveKeysounds(.Lane + 128) = .Value
                                    If .Lane = 7 Then
                                        ActiveKeysounds(8 + 128) = .Value
                                    End If
                                Case 4
                                    .Parameter = ChartParameter.TempoChange
                                    .Value = Val3 + (Val2 * 256)
                                    If ThisBPM = 0 Then
                                        ThisBPM = .Value
                                    End If
                                    .Lane = 1
                                Case 5
                                    .Parameter = ChartParameter.Meter
                                    .Value = Val3
                                    .Lane = 4
                                Case 6
                                    If Val2 = 0 Then
                                        .Parameter = ChartParameter.EndSong
                                        Chart.SongLengthSeconds = (.OffsetBase * TickRate) \ 1000
                                    End If
                                Case 7
                                    .Parameter = ChartParameter.BGM
                                    .Value = Val3
                                    If .Value = 1 Then
                                        If BGMused Then
                                            ValidNote = False
                                        Else
                                            BGMused = True
                                        End If
                                    End If
                                Case 8
                                    .Parameter = ChartParameter.Judgement
                                    .Value = Val3
                                    .Lane = Val2
                                Case 12
                                    .Parameter = ChartParameter.Measure
                                    .Player = Val2
                                    If (.OffsetBase <= 0) Or (.Player <> 0) Then
                                        ValidNote = False
                                    End If
                                Case Else
                                    ValidNote = False
                            End Select
                            If .OffsetBase = &H7FFF Then
                                ValidNote = False
                                .Parameter = ChartParameter.Bad
                            End If
                        End With
                        ChartOffset += 4
                        If ValidNote Then
                            .EventCount += 1
                        End If
                    Loop
                End With


            Case ChartTypes.IIDXCS '==========================================
                Chart.HasProgramChanges = True
                ConvertToMetric = True
                If ChartData(0) <> 8 Then
                    DebugLog("ChartLoad", "Chart signature is invalid.", "IIDXCS")
                    Return True
                    Exit Function
                End If
                TickRate = DataMakeInt32(ChartData(4), ChartData(5), ChartData(6), ChartData(7))
                ChartOffset = 8
                'bypass note counts...
                If ChartData(ChartOffset + 4) = 0 Then
                    Do
                        If (ChartData(ChartOffset + 5) < 250) Then
                            ChartOffset += 8
                            Exit Do
                        End If
                        ChartOffset += 8
                    Loop While ChartData(ChartOffset + 4) = 0
                End If
                If ChartData(ChartOffset + 4) = 1 Then
                    Do
                        If (ChartData(ChartOffset + 5) < 250) Then
                            ChartOffset += 8
                            Exit Do
                        End If
                        ChartOffset += 8
                    Loop While ChartData(ChartOffset + 4) = 1
                End If
                e = False
                With Chart
                    Do While (ChartOffset < UBound(ChartData)) And (Not e)
                        ValidNote = True
                        ReDim Preserve .Events(.EventCount)
                        With .Events(.EventCount)
                            .BMSUsed = False
                            .OffsetBase = DataMakeInt32(ChartData(ChartOffset), ChartData(ChartOffset + 1), ChartData(ChartOffset + 2), ChartData(ChartOffset + 3))
                            If .OffsetBase < ReferenceBase Then
                                'backwards step
                            Else
                                ReferenceBase = .OffsetBase
                            End If
                            Val1 = ChartData(ChartOffset + 4) And &HF
                            Val2 = ChartData(ChartOffset + 4) \ 16
                            Val3 = DataMakeInt32(ChartData(ChartOffset + 6), ChartData(ChartOffset + 7))
                            If ChartData(ChartOffset + 5) <> 0 And Val3 = 0 Then
                                Val3 = ChartData(ChartOffset + 5)
                            End If
                            Select Case Val1
                                Case 0
                                    .Player = 1
                                    .Lane = Val2
                                    .Parameter = ChartParameter.Note
                                    If .OffsetBase = 0 Then
                                        ValidNote = False
                                    End If
                                    .Value = ActiveKeysounds(.Lane)
                                    If .Value = 0 Then
                                        ValidNote = False
                                    End If
                                Case 1
                                    .Player = 2
                                    .Lane = Val2
                                    .Parameter = ChartParameter.Note
                                    If .OffsetBase = 0 Then
                                        ValidNote = False
                                    End If
                                    .Value = ActiveKeysounds(.Lane + 128)
                                    If .Value = 0 Then
                                        ValidNote = False
                                    End If
                                Case 2
                                    .Player = 1
                                    .Parameter = ChartParameter.SoundChange
                                    .Value = Val3
                                    .Lane = Val2
                                    ActiveKeysounds(.Lane) = .Value
                                Case 3
                                    .Player = 2
                                    .Parameter = ChartParameter.SoundChange
                                    .Value = Val3
                                    .Lane = Val2
                                    ActiveKeysounds(.Lane + 128) = .Value
                                Case 4
                                    .Parameter = ChartParameter.TempoChange
                                    .Value = Val3 Or (Val2 << 8)
                                    If ThisBPM = 0 Then
                                        ThisBPM = .Value
                                    End If
                                    .Lane = 1
                                Case 5
                                    .Parameter = ChartParameter.Meter
                                    .Value = Val3
                                    .Lane = 4
                                Case 6
                                    If Val2 = 0 Then
                                        .Parameter = ChartParameter.EndSong
                                        Chart.SongLengthSeconds = (.OffsetBase * TickRate) \ 1000
                                        e = True
                                    End If
                                Case 7
                                    .Parameter = ChartParameter.BGM
                                    .Value = Val3
                                Case 8
                                    .Parameter = ChartParameter.Judgement
                                    .Value = Val3
                                    .Lane = Val2
                                Case 12
                                    .Parameter = ChartParameter.Measure
                                    .Player = Val2
                                    If (.OffsetBase <= 0) Or (.Player <> 0) Then
                                        ValidNote = False
                                    End If
                                Case Else
                                    ValidNote = False
                            End Select
                        End With
                        ChartOffset += 8
                        If ValidNote Then
                            .EventCount += 1
                        End If
                    Loop
                End With

            Case ChartTypes.IIDXAC
                If iSubChart < 0 Or iSubChart >= 12 Then
                    'bad subchart number
                    Return True
                    Exit Function
                End If
                x = iSubChart * 8
                ChartOffset = DataMakeInt32(ChartData(x), ChartData(x + 1), ChartData(x + 2), ChartData(x + 3))
                ChartSize = DataMakeInt32(ChartData(x + 4), ChartData(x + 5), ChartData(x + 6), ChartData(x + 7))
                If ChartSize <= 0 Then
                    'zero length
                    Return True
                    Exit Function
                End If
                If ChartOffset <= 0 Then
                    'zero offset
                    Return True
                    Exit Function
                End If
                e = False
                With Chart
                    Do While (ChartOffset < ChartSize) And (Not e)
                        ValidNote = True
                        ReDim Preserve .Events(.EventCount)
                        With .Events(.EventCount)
                            .BMSUsed = False
                            .OffsetBase = DataMakeInt32(ChartData(ChartOffset), ChartData(ChartOffset + 1), ChartData(ChartOffset + 2), ChartData(ChartOffset + 3))
                            If .OffsetBase < ReferenceBase Then
                                'backwards step
                            Else
                                ReferenceBase = .OffsetBase
                            End If
                            Val1 = ChartData(ChartOffset + 4)
                            Val2 = ChartData(ChartOffset + 5)
                            Val3 = DataMakeInt32(ChartData(ChartOffset + 6), ChartData(ChartOffset + 7))
                            Select Case Val1
                                Case 0
                                    .Player = 1
                                    .Lane = Val2
                                    .Parameter = ChartParameter.Note
                                    If .OffsetBase = 0 Then
                                        ValidNote = False
                                    End If
                                    .Value = ActiveKeysounds(.Lane)
                                    If .Value = 0 Then
                                        ValidNote = False
                                    End If
                                    If Val3 = 0 Then
                                        LastNoteOffset(.Lane) = Chart.EventCount
                                    ElseIf ValidNote Then
                                        If LastNoteOffset(.Lane) > 0 Then
                                            Chart.Events(LastNoteOffset(.Lane)).Parameter = ChartParameter.FreezeNote
                                            .Parameter = ChartParameter.FreezeNote
                                            LastNoteOffset(.Lane) = 0
                                        End If
                                    End If
                                Case 1
                                    .Player = 2
                                    .Lane = Val2
                                    .Parameter = ChartParameter.Note
                                    If .OffsetBase = 0 Then
                                        ValidNote = False
                                    End If
                                    .Value = ActiveKeysounds(.Lane + 128)
                                    If .Value = 0 Then
                                        ValidNote = False
                                    End If
                                    If Val3 = 0 Then
                                        LastNoteOffset(.Lane + 128) = Chart.EventCount
                                    ElseIf ValidNote Then
                                        If LastNoteOffset(.Lane + 128) > 0 Then
                                            Chart.Events(LastNoteOffset(.Lane + 128)).Parameter = ChartParameter.FreezeNote
                                            .Parameter = ChartParameter.FreezeNote
                                            LastNoteOffset(.Lane + 128) = 0
                                        End If
                                    End If
                                Case 2
                                    .Player = 1
                                    .Parameter = ChartParameter.SoundChange
                                    .Value = Val3
                                    .Lane = Val2
                                    ActiveKeysounds(.Lane) = .Value
                                Case 3
                                    .Player = 2
                                    .Parameter = ChartParameter.SoundChange
                                    .Value = Val3
                                    .Lane = Val2
                                    ActiveKeysounds(.Lane + 128) = .Value
                                Case 4
                                    .Parameter = ChartParameter.TempoChange
                                    If Val2 > 0 Then
                                        .Value = Val3 / Val2
                                    Else
                                        .Value = Val3
                                    End If
                                    If ThisBPM = 0 Then
                                        ThisBPM = .Value
                                    End If
                                    .Lane = 1
                                Case 5
                                    .Parameter = ChartParameter.Meter
                                    .Value = Val3
                                    .Lane = 4
                                Case 6
                                    If Val2 = 0 Then
                                        .Parameter = ChartParameter.EndSong
                                        Chart.SongLengthSeconds = (.OffsetBase * TickRate) \ 1000
                                        e = True
                                    End If
                                Case 7
                                    .Parameter = ChartParameter.BGM
                                    .Value = Val3
                                Case 8
                                    .Parameter = ChartParameter.Judgement
                                    .Value = Val3
                                    .Lane = Val2
                                Case 12
                                    .Parameter = ChartParameter.Measure
                                    .Player = Val2
                                    If (.OffsetBase <= 0) Or (.Player <> 0) Then
                                        ValidNote = False
                                    End If
                                Case Else
                                    ValidNote = False
                            End Select
                        End With
                        ChartOffset += 8
                        If ValidNote Then
                            .EventCount += 1
                        End If
                    Loop
                End With

            Case Else
                'not supported
                DebugLog("ChartLoad", "Chart type is not supported.", "Type: " & iChartType)
                Return True
                Exit Function
        End Select

        Chart.ChartType = iChartType
        If ChartForceTiming > 0 Then
            TickRate = ChartForceTiming
        End If

        'METRICS STAGE ***********************************************************
        If ThisBPM = 0 Then
            'requires a BPM be set to start
            DebugLog("ChartLoad", "Conversion to metric offsets can't be done.", "There is no BPM.")
            Return True
            Exit Function
        End If

        If Not ConvertToMetric Then
            'sort by metric offset, then by parameter priority
            Do
                e = False
                For x = 1 To Chart.EventCount - 1
                    If (Chart.Events(x).OffsetMetric < Chart.Events(x - 1).OffsetMetric) Or ((Chart.Events(x).OffsetMetric = Chart.Events(x - 1).OffsetMetric) And (Chart.Events(x).Parameter < Chart.Events(x - 1).Parameter)) Then
                        TempEvent = Chart.Events(x)
                        Chart.Events(x) = Chart.Events(x - 1)
                        Chart.Events(x - 1) = TempEvent
                        e = True
                    End If
                Next
            Loop While e
            'now calculate offsets in milliseconds
            LastMeasureOffset = 0
            ThisMetric = 0
            ThisMeasureDefaultSize = ((ThisBPM / 4) / 60000)
            ReferenceBase = 0
            For x = 0 To Chart.EventCount - 1
                With Chart.Events(x)
                    If .Parameter = ChartParameter.EndSong Then
                        For y = LastMeasureOffset To x
                            Chart.Events(y).OffsetMSec = ReferenceBase + ((Chart.Events(y).OffsetMetric - ThisMetric) / ThisMeasureDefaultSize)
                            Chart.Events(y).OffsetBase = Chart.Events(y).OffsetMSec
                        Next
                        Chart.EventCount = x + 1
                        Exit For
                    ElseIf .Parameter = ChartParameter.TempoChange Then
                        For y = LastMeasureOffset To x
                            Chart.Events(y).OffsetMSec = ReferenceBase + ((Chart.Events(y).OffsetMetric - ThisMetric) / ThisMeasureDefaultSize)
                            Chart.Events(y).OffsetBase = Chart.Events(y).OffsetMSec
                        Next
                        ThisMetric = Chart.Events(x).OffsetMetric
                        LastMeasureOffset = x + 1
                        ReferenceBase = .OffsetMSec
                        ThisBPM = .Value / .Lane
                        ThisMeasureDefaultSize = ((ThisBPM / 4) / 60000)
                    End If
                End With
            Next
            x = x
        Else

            'sort by base offset, then by parameter priority
            Do
                e = False
                For x = 1 To Chart.EventCount - 1
                    If (Chart.Events(x).OffsetBase < Chart.Events(x - 1).OffsetBase) Or ((Chart.Events(x).OffsetBase = Chart.Events(x - 1).OffsetBase) And (Chart.Events(x).Parameter < Chart.Events(x - 1).Parameter)) Then
                        TempEvent = Chart.Events(x)
                        Chart.Events(x) = Chart.Events(x - 1)
                        Chart.Events(x - 1) = TempEvent
                        e = True
                    End If
                Next
            Loop While e

            'create a list of the measure offsets
            ReDim Chart.MeasureOffsets(0)
            Chart.MeasureOffsets(0) = 0
            MeasureCount = 1
            For x = 0 To Chart.EventCount - 1
                With Chart.Events(x)
                    If (.Parameter = ChartParameter.Measure Or .Parameter = ChartParameter.EndSong) Then
                        If .OffsetBase > Chart.MeasureOffsets(MeasureCount - 1) Then
                            ReDim Preserve Chart.MeasureOffsets(MeasureCount)
                            Chart.MeasureOffsets(MeasureCount) = .OffsetBase
                            MeasureCount += 1
                        End If
                    End If
                End With
            Next
            ReDim Preserve Chart.MeasureOffsets(MeasureCount)

            'if we force timing, do it here
            If iForceTiming > 0 Then
                TickRate = iForceTiming
            End If

            'calculate offsets per event in milliseconds
            For x = 0 To Chart.EventCount - 1
                Chart.Events(x).OffsetMSec = (TickRate * Chart.Events(x).OffsetBase) \ 1000
            Next

            'calculate metric offsets per event
            ThisMetric = 0
            ThisMeasure = 0
            ReferenceBase = 0
            LastMeasureOffset = 0
            ReDim MeasureMetrics(MeasureCount)
            ReDim MeasureMetricOffsets(MeasureCount)
            ReDim Chart.MeasurePreferredUnits(MeasureCount)
            ThisMeasureDefaultSize = ChartNormalMeasureSize(TickRate, ThisBPM)
            ThisMetric = 0
            HasBPMChange = False
            For x = 0 To Chart.EventCount - 1
                With Chart.Events(x)
                    .OffsetMetric = ThisMetric + ((.OffsetBase - ReferenceBase) / ThisMeasureDefaultSize)
                    .Measure = ThisMeasure
                    Select Case .Parameter
                        Case ChartParameter.EndSong, ChartParameter.Measure
                            If .OffsetBase > LastMeasureOffset Then
                                MeasureMetrics(ThisMeasure) += ((.OffsetBase - LastMeasureOffset) / ThisMeasureDefaultSize)
                                If (Not HasBPMChange) And (TickRate > 1000) Then
                                    Chart.MeasurePreferredUnits(ThisMeasure) = .OffsetBase - LastMeasureOffset
                                End If
                                ThisMeasure += 1
                                LastMeasureOffset = .OffsetBase
                                MeasureMetricOffsets(ThisMeasure) = .OffsetMetric
                                HasBPMChange = False
                            End If
                        Case ChartParameter.TempoChange
                            If .OffsetBase > 0 Then
                                If (.OffsetBase <> LastMeasureOffset) Then
                                    HasBPMChange = True
                                End If
                                MeasureMetrics(ThisMeasure) += ((.OffsetBase - LastMeasureOffset) / ThisMeasureDefaultSize)
                                ThisMetric = .OffsetMetric
                                ReferenceBase = .OffsetBase
                                LastMeasureOffset = .OffsetBase
                                ThisBPM = .Value / .Lane
                                ThisMeasureDefaultSize = ChartNormalMeasureSize(TickRate, ThisBPM)
                            End If
                    End Select
                End With
            Next
            For x = 0 To Chart.EventCount - 1
                With Chart.Events(x)
                    If .Measure < UBound(MeasureMetricOffsets) Then
                        .OffsetWithinMeasure = (.OffsetMetric - MeasureMetricOffsets(.Measure)) / (MeasureMetricOffsets(.Measure + 1) - MeasureMetricOffsets(.Measure))
                        .OffsetBaseWithinMeasure = (.OffsetBase - Chart.MeasureOffsets(.Measure))
                        'clamp offset
                        If (.OffsetWithinMeasure < 0) Then
                            .OffsetWithinMeasure = 0
                        ElseIf (.OffsetWithinMeasure > 1) Then
                            .OffsetWithinMeasure = 1
                        End If
                    End If
                End With
            Next

            For x = 0 To UBound(MeasureMetrics)
                If MeasureMetrics(x) = 0 Then MeasureMetrics(x) = 1
            Next
        End If

        If Not HasWaveTable Then
            For x = 0 To Chart.EventCount - 1
                With Chart.Events(x)
                    If (.Parameter = ChartParameter.Note) And (.Value < UBound(Chart.WaveTable)) And (.Value > 0) Then
                        Chart.WaveTable(.Value) = DataBMEString(.Value) & ".wav"
                    End If
                End With
            Next
        End If

        Chart.MeasureMetrics = MeasureMetrics

        Return False
    End Function

    Public Sub ChartSaveBMS(ByVal sFileName As String, ByVal sBMSType As ChartTypes, ByVal iPrecision As Integer, Optional ByVal sHeaderString As String = "", Optional ByVal sWave01 As String = "", Optional ByVal sWavePrefix As String = "", Optional ByVal iPlayLevel As Integer = 0)
        Dim HighestWave As Integer = -1
        Dim WaveList() As Boolean
        Dim WaveMap() As Integer
        Dim BPMMap() As Double
        Dim NoteQuery() As ChartFormat
        Dim ThisMap As Integer = 1
        Dim Writer As IO.StreamWriter
        Dim ThisMeasurePrecision As Integer
        Dim UseMeasurePrecision As Boolean
        Dim EmptyLine As String
        ReDim WaveList(0)
        ReDim BPMMap(0)
        ReDim NoteQuery(0)
        Dim MeasureCarry As Double = 0
        Dim ThisMeasureFrac As Double = 0
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim a As Integer
        Dim s As String
        Dim t As String
        Dim u As String
        Dim v As String
        Dim w As String
        Dim b As Integer
        Dim e As Boolean
        Dim bReduce As Boolean
        For x = 0 To UBound(Chart.Events)
            With Chart.Events(x)
                .BMSUsed = False
                If (.Parameter = ChartParameter.SoundChange) Or (.Parameter = ChartParameter.BGM) Then
                    If .Value > HighestWave Then
                        HighestWave = .Value
                        ReDim Preserve WaveList(HighestWave)
                    End If
                    WaveList(.Value) = True
                ElseIf (.Parameter = ChartParameter.TempoChange) Then
                    If (.OffsetBase = 0) Then
                        BPMMap(0) = (.Value / .Lane)
                    ElseIf ((.Value / .Lane) > 255) Or ((.Value / .Lane) <> Int(.Value / .Lane)) Then
                        Do
                            For y = 1 To ThisMap - 1
                                If BPMMap(y) = (.Value / .Lane) Then
                                    Exit Do
                                End If
                            Next
                            ReDim Preserve BPMMap(ThisMap)
                            BPMMap(ThisMap) = (.Value / .Lane)
                            ThisMap += 1
                        Loop Until (True = True)
                    End If
                End If
            End With
        Next
        If HighestWave <= 0 Then
            Exit Sub
        End If
        ThisMap = 1
        ReDim WaveMap(HighestWave)
        For x = 1 To HighestWave
            If WaveList(x) Then
                WaveMap(x) = ThisMap
                ThisMap += 1
            End If
        Next

        Writer = New IO.StreamWriter(sFileName)
        Writer.WriteLine(";BMR(ESE) ver " & Application.ProductVersion)
        Writer.WriteLine("; notes: " & Chart.NoteCount(1) & "+" & Chart.NoteCount(2) & "=" & Chart.NoteCount)
        Writer.WriteLine()
        If sHeaderString <> "" Then
            Writer.WriteLine(sHeaderString)
        End If
        If Chart.NoteCount(2) > 0 Then
            Writer.WriteLine("#PLAYER 3")
        Else
            Writer.WriteLine("#PLAYER 1")
        End If
        Writer.WriteLine("#BPM " & BPMMap(0))
        For x = 1 To UBound(BPMMap)
            Writer.WriteLine("#BPM" & DataBMEString(x, 2, "0123456789") & " " & BPMMap(x))
        Next
        If Chart.NoteCount > 0 Then
            Writer.WriteLine("#TOTAL " & CStr(Int(Chart.NoteCount * (3072.0 / Chart.NoteCount) / 1024.0 * 100.0)))
        End If
        If iPlayLevel > 0 Then
            Writer.WriteLine("#PLAYLEVEL " & CStr(iPlayLevel))
        End If
        If sWave01 <> "" Then
            z = 2
            Writer.WriteLine("#WAV" & DataBMEString(WaveMap(1)) & " " & sWave01)
        Else
            z = 1
        End If
        For x = z To HighestWave
            If WaveMap(x) > 0 And WaveMap(x) < 1295 Then
                Writer.WriteLine("#WAV" & DataBMEString(WaveMap(x)) & " " & sWavePrefix & DataKeySoundFileName(x) & ".wav")
            End If
        Next

        'do some selective rounding - not all sims calculate measure sizes in
        'double floating-point precision but almost all certainly in at least single precision,
        'so round to 5 decimal places for better accuracy in those sims and carry the remainder
        'to subsequent measure lengths
        For x = 0 To UBound(Chart.MeasureMetrics)
            ThisMeasureFrac = Math.Round(Chart.MeasureMetrics(x) + MeasureCarry, 5)
            MeasureCarry -= ThisMeasureFrac - Chart.MeasureMetrics(x)
            Writer.WriteLine("#" & DataPadString(CStr(x), 3, "0") & "02:" & CStr(ThisMeasureFrac))
        Next

        u = "00"
        For x = 0 To UBound(Chart.MeasureMetrics)
            s = "#" & DataPadString(CStr(x), 3, "0")
            For z = 0 To 999
                UseMeasurePrecision = False
                ThisMeasurePrecision = Math.Max(iPrecision, iPrecision * Int(Chart.MeasureMetrics(x)))
                EmptyLine = Strings.StrDup(ThisMeasurePrecision * 2, "0")
                Select Case Chart.ChartType
                    Case ChartTypes.Popn4, ChartTypes.Popn6, ChartTypes.Popn8
                        Select Case z
                            Case 0 : e = Chart.SelectNotes(NoteQuery, ChartParameter.BGM, x) : u = "01"
                            Case 1 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 0, 1) : u = "11"
                            Case 2 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 1, 1) : u = "12"
                            Case 3 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 2, 1) : u = "13"
                            Case 4 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 3, 1) : u = "14"
                            Case 5 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 4, 1) : u = "15"
                            Case 6 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 5, 1) : u = "22"
                            Case 7 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 6, 1) : u = "23"
                            Case 8 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 7, 1) : u = "24"
                            Case 9 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 8, 1) : u = "25"
                            Case 10 : e = Chart.SelectNotes(NoteQuery, ChartParameter.TempoChange, x) : u = "03" : ThisMeasurePrecision *= 2
                            Case 11 : e = Chart.SelectNotes(NoteQuery, ChartParameter.TempoChange, x) : u = "08" : ThisMeasurePrecision *= 2
                            Case Else
                                Exit For
                        End Select
                    Case ChartTypes.IIDXCS, ChartTypes.IIDXCS2, ChartTypes.IIDXAC, ChartTypes.IIDXCS5
                        Select Case z
                            Case 0 : e = Chart.SelectNotes(NoteQuery, ChartParameter.BGM, x) : u = "01"
                            Case 1 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 0, 1) : u = "11"
                            Case 2 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 1, 1) : u = "12"
                            Case 3 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 2, 1) : u = "13"
                            Case 4 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 3, 1) : u = "14"
                            Case 5 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 4, 1) : u = "15"
                            Case 6 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 5, 1) : u = "18"
                            Case 7 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 6, 1) : u = "19"
                            Case 8 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 7, 1) : u = "16"
                            Case 9 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 8, 1) : u = "17"
                            Case 10 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 0, 2) : u = "21"
                            Case 11 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 1, 2) : u = "22"
                            Case 12 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 2, 2) : u = "23"
                            Case 13 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 3, 2) : u = "24"
                            Case 14 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 4, 2) : u = "25"
                            Case 15 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 5, 2) : u = "28"
                            Case 16 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 6, 2) : u = "29"
                            Case 17 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 7, 2) : u = "26"
                            Case 18 : e = Chart.SelectNotes(NoteQuery, ChartParameter.Note, x, 8, 2) : u = "27"
                            Case 19 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 0, 1) : u = "51"
                            Case 20 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 1, 1) : u = "52"
                            Case 21 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 2, 1) : u = "53"
                            Case 22 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 3, 1) : u = "54"
                            Case 23 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 4, 1) : u = "55"
                            Case 24 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 5, 1) : u = "58"
                            Case 25 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 6, 1) : u = "59"
                            Case 26 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 7, 1) : u = "56"
                            Case 27 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 0, 2) : u = "61"
                            Case 28 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 1, 2) : u = "62"
                            Case 29 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 2, 2) : u = "63"
                            Case 30 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 3, 2) : u = "64"
                            Case 31 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 4, 2) : u = "65"
                            Case 32 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 5, 2) : u = "68"
                            Case 33 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 6, 2) : u = "69"
                            Case 34 : e = Chart.SelectNotes(NoteQuery, ChartParameter.FreezeNote, x, 7, 2) : u = "66"
                            Case 35 : e = Chart.SelectNotes(NoteQuery, ChartParameter.TempoChange, x) : u = "03" : ThisMeasurePrecision *= 2
                            Case 36 : e = Chart.SelectNotes(NoteQuery, ChartParameter.TempoChange, x) : u = "08" : ThisMeasurePrecision *= 2
                            Case Else
                                Exit For
                        End Select
                End Select
                If Chart.MeasurePreferredUnits(x) > 0 Then
                    ThisMeasurePrecision = Chart.MeasurePreferredUnits(x)
                    UseMeasurePrecision = True
                End If

                If e Then
                    e = True
                    Do While e = True
                        t = Strings.StrDup(ThisMeasurePrecision * 2, "0")
                        e = False
                        For a = 0 To UBound(NoteQuery)
                            With NoteQuery(a)
                                If (UseMeasurePrecision = True) Or (.OffsetWithinMeasure >= 0 And .OffsetWithinMeasure <= 1) Then
                                    If (UseMeasurePrecision) Then
                                        b = .OffsetBase - Chart.MeasureOffsets(x)
                                    Else
                                        b = Int((.OffsetWithinMeasure * ThisMeasurePrecision) + 0.5)
                                    End If
                                    If b >= ThisMeasurePrecision Then
                                        b = ThisMeasurePrecision - 1
                                    ElseIf b < 0 Then
                                        b = 0
                                    End If
                                    b *= 2
                                    If Not .BMSUsed Then
                                        Select Case .Parameter
                                            Case ChartParameter.Note
                                                If .Value > 0 Then
                                                    Mid(t, b + 1, 2) = DataBMEString(WaveMap(.Value))
                                                End If
                                                .BMSUsed = True
                                            Case ChartParameter.BGM
                                                If Mid(t, b + 1, 2) = "00" Then
                                                    Mid(t, b + 1, 2) = DataBMEString(WaveMap(.Value))
                                                    .BMSUsed = True
                                                Else
                                                    e = True
                                                End If
                                            Case ChartParameter.TempoChange
                                                If .OffsetBase > 0 Then
                                                    If u = "08" Then
                                                        For y = 1 To UBound(BPMMap)
                                                            If BPMMap(y) = (.Value / .Lane) Then
                                                                .BMSUsed = True
                                                                Mid(t, b + 1, 2) = DataBMEString(y, 2, "0123456789")
                                                                Exit Select
                                                            End If
                                                        Next
                                                        'didn't find it in bpm map
                                                    ElseIf u = "03" Then
                                                        .BMSUsed = True
                                                        Mid(t, b + 1, 2) = DataBMEString(Int(.Value / .Lane), 2, "0123456789ABCDEF")
                                                    End If
                                                End If
                                        End Select
                                    End If
                                End If
                            End With
                        Next
                        If u <> "00" Then
                            If t <> StrDup(Len(t), "0") Then
                                'line reduction
                                Do
                                    bReduce = False
                                    For y = 2 To (Len(t) \ 2)
                                        w = StrDup((y - 1) * 2, "0")
                                        bReduce = True
                                        If (Len(t) \ y) = (Len(t) / y) Then
                                            For a = 1 To Len(t) Step (y * 2)
                                                If Strings.Mid(t, a + 2, ((y - 1) * 2)) <> w Then
                                                    bReduce = False
                                                    Exit For
                                                End If
                                            Next
                                        Else
                                            bReduce = False
                                        End If
                                        If bReduce Then
                                            v = ""
                                            For a = 1 To Len(t) Step (y * 2)
                                                v &= Strings.Mid(t, a, 2)
                                            Next
                                            t = v
                                            Exit For
                                        End If
                                    Next
                                Loop While bReduce = True
                                Writer.WriteLine(s & u & ":" & t)
                            End If
                        End If
                    Loop
                End If
            Next
        Next
        Writer.Close()

    End Sub

    Public Function ChartNormalMeasureSize(ByVal TickRate As Integer, ByVal BPM As Double) As Double
        Return (((CDbl(60) / BPM) * CDbl(4)) * (CDbl(1000000) / CDbl(TickRate)))
    End Function

End Module
