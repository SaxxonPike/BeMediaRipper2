Option Explicit On
Imports System.Runtime.InteropServices.Marshal
Module modData

    Public SongString As String
    Public sSourceFile As String = ""
    Public sTargetFolder As String = ""
    Public sExtraFile As String = ""
    Public iGameType As Integer = 0
    Public DataNamingStyle As Integer = 0
    'Public DataRipCharts As Boolean = False
    'Public DataConvertCharts As Boolean = False
    'Public DataRipSounds As Boolean = False
    'Public DataConvertSounds As Boolean = False
    Public bChartInfo As Boolean
    Public bOneClickMode As Boolean

    Private sBadCharacters As String = "*|:""<>?/\"
    Private DJMainVolumeTable(0 To 255) As Double
    Private FirebeatVolumeTable(0 To 255) As Double
    Private PanTableLeft(0 To &H100) As Double
    Private PanTableRight(0 To &H100) As Double
    Private LastDJMainOffset As Integer

    Private LogWrite As System.IO.StreamWriter

    Public Structure DataSampleHeaderA11
        Public Ident As Short
    End Structure
    Public Structure DataSampleHeaderB11
        Public SampleCount As Integer
        Public TotalLength As Integer
    End Structure
    Public Structure DataSampleInfoA11
        Public Unk0 As Short
        Public unk1 As Byte
        Public ChanCount As Byte
        Public Unk2 As Integer
        Public PanLeft As Byte
        Public PanRight As Byte
        Public SampleNum As Short
        Public volume As Byte
        Public unk3 As Byte
        Public Unk4 As Short
    End Structure
    Public Structure DataSampleInfoB11
        Public SampOffset As Integer
        Public SampLength As Integer
        Public ChanCount As Short
        Public Frequ As Short
        Public Unk0 As Integer
    End Structure
    Public Structure DataSampleInfo3
        Public SampleNum As Short
        Public Unk0 As Short
        Public unk1 As Byte
        Public vol As Byte
        Public pan As Byte
        Public SampType As Byte
        Public FreqLeft As Integer
        Public FreqRight As Integer
        Public OffsLeft As Integer
        Public OffsRight As Integer
        Public PseudoLeft As Integer
        Public PseudoRight As Integer
    End Structure
    Private Enum DataCHDType As Integer
        None
        DJMain
        Firebeat
    End Enum
    Private Structure DataCHDOffsets
        Public PlaybackChannel As Integer
        Public Frequency As Integer
        Public Pan As Integer
        Public Volume As Integer
        Public Offset As Integer
        Public Length As Integer
        Public SampleType As Integer
        Public Flags As Integer
    End Structure
    Public Structure DataConvertedSample
        Public Sample() As Short
        Public VolumeL As Double
        Public VolumeR As Double
        Public ReadOnly Property Length() As Integer
            Get
                Return UBound(Sample)
            End Get
        End Property
    End Structure
    Public Enum DataEncodingType As Integer
        None
        KonamiLZ77
        KonamiLZSS
    End Enum
    Public Enum DataDetectedType As Integer
        None
        DDRPSX
        DDRPS2
        IIDXKeysound
        IIDXOldChart
        IIDXNewChart
        IIDXOldKeysound
        IIDXOldBGM
        BeatmaniaChart
        BeatmaniaKeysound
        BeatmaniaBGM
        VOB
        FirebeatChart
        FirebeatBGM
        FirebeatKeysound
        DJmainChart
        DJmainBGM
        DJmainKeysound
        VAG
    End Enum

    Public Structure DataFileEntry
        Public Name As String
        Public Offset As Long
        Public Length As Long
        Public DecodePath As String
        Public DecodeType As SoundDecodeType
        Public IsFromInfoFile As Boolean
        Public ForcedType As DataDetectedType
        Public Keysound As Integer
        Public BGM As Integer
        Public Movie As Integer
        Public Prefix As String
        Public Suffix As String
        Public Difficulty As Integer
        Public ForceTiming As Integer
        Public Artist As String
        Public Genre As String
        Public ReadOnly Property FullName() As String
            Get
                Return Prefix & Name & Suffix
            End Get
        End Property
    End Structure

    Public Structure DataSongInfoType
        Public Name As String
        Public Charts() As Integer
        Public Keysounds() As Integer
        Public BGM() As Integer
        Public ChartNames() As String
        Public ChartDifficulty() As Integer
        Public Movies() As Integer
    End Structure

    Public DataFiles() As DataFileEntry
    Public DataFileCount As Integer
    Public Songs() As DataSongInfoType
    Public SongsCount As Integer
    Public DataFileSystem As Integer
    Public DataSongSystem As Integer
    Public DataConfig As Integer

    Public Sub DataInit()
        Dim x As Integer
        For x = 0 To 255
            DJMainVolumeTable(x) = Math.Pow(10.0, (-36.0 * x / 64) / 20.0) * 0.8
            FirebeatVolumeTable(x) = Math.Pow(10.0, (-36.0 * x / 144) / 20.0) * 0.8
        Next
        For x = 0 To &H100
            PanTableLeft(x) = (Math.Sqrt(x) / Math.Sqrt(&H100))
        Next
        For x = 0 To &H100
            PanTableRight(x) = PanTableLeft(&H100 - x)
        Next
    End Sub


    Public Function DataStringToBytes(ByVal str As String) As Byte()
        Dim enc As New System.Text.ASCIIEncoding()
        Return enc.GetBytes(str)
    End Function

    Public Function DataBytesToString(ByVal dBytes As Byte()) As String
        Dim enc As New System.Text.ASCIIEncoding()
        Return enc.GetString(dBytes)
    End Function

    Public Function DataPadString(ByVal sString As String, ByVal iMinSize As Integer, ByVal cPadChar As String) As String
        DataPadString = sString
        Do While Len(DataPadString) < iMinSize
            DataPadString = cPadChar & DataPadString
        Loop
    End Function

    Public Function DataGetFormattedName(ByVal sSongName As String) As String
        Dim sn As String
        Dim x As Integer
        'DataGetFormattedName = ""
        'If sSongNumber <= UBound(Songs) Then
        DataGetFormattedName = ""
        If sSongName <> "" Then
            sn = DataTranslateName(sSongName)
            sn = Trim(sn)
            For x = 1 To Len(sBadCharacters)
                sn = Replace(sn, Mid(sBadCharacters, x, 1), "_")
            Next
            Do While Right(sn, 1) = "."
                sn = Strings.Left(sn, Len(sn) - 1)
            Loop
            Do
                x = InStr(sn, Chr(&H81))
                If x > 0 Then
                    sn = Strings.Left(sn, x - 1) & Strings.Mid(sn, x + 2)
                Else
                    Exit Do
                End If
            Loop
            DataGetFormattedName = sn
        End If
        '    End If
    End Function

    Public Function DataTranslateName(ByVal sSongName As String) As String
        Dim sn As String = sSongName
        sn = Replace(sn, Chr(&H81) & Chr(&H43), ",")
        sn = Replace(sn, Chr(&H81) & Chr(&H60), "~")
        sn = Replace(sn, Chr(&H81) & Chr(&H9A), "*")
        Return sn
    End Function

    Public Function DataReadFileTable(ByVal FileName As String, ByVal ConfigToUse As Integer) As Boolean
        Dim f As Integer = FreeFile()
        Dim b As Byte = 0
        Dim i As Short = 0
        Dim l As Integer = 0
        Dim o As Long = 0
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim a As String = ""
        Dim s As String = ""
        Dim ExecutableOffsets() As Integer = {0}
        Dim ExecutableNames() As String = {""}
        Dim ExecutableIndex As Integer = -1
        Dim d1 As Byte = 0
        Dim d2 As Short = 0
        Dim d4 As Integer = 0
        Dim db() As Byte
        Dim ChartAlloc() As Integer
        Dim ChartName() As String
        Dim ChartName5() As String
        Dim DiffIndex() As Integer
        Dim DiffList() As Integer
        Dim KeysList() As Integer
        Dim b5Key As Boolean
        Dim SongName As String = ""
        Dim SongNameList() As String = {""}
        Dim SongNameListCount As Integer = 0
        Dim fs As Integer = 0
        Dim gs As Integer = 0
        Dim offs As Integer = 0
        Dim leng As Integer = 0
        Dim lkey As Integer
        Dim lbgm As Integer
        Dim lmov As Integer
        Dim Offset As Long = 0
        Dim Length As Long = 0
        Dim FileFormat As Integer = -1
        Dim GameFormat As Integer = -1
        Dim ChartTime As Integer
        Dim LastOffset As Long = 0
        Dim flog As Integer
        Dim ind As Integer = 0
        Dim bValid As Boolean
        DataFileCount = 0
        SongsCount = 0
        ReDim DataFiles(0)
        ReDim db(0)
        ReDim Songs(0)
        ReDim ChartName(0)
        ReDim ChartName5(0)
        ReDim ChartAlloc(0)
        ReDim DiffIndex(0)
        ReDim DiffList(0)
        ReDim KeysList(0)
        Dim SongDBInfo As ConfigSongDBType = SongDB(0)
        DataReadFileTable = False

        bChartInfo = False
        FileFormat = Config(ConfigToUse).TableType
        GameFormat = Config(ConfigToUse).SongType

        'format search
        fs = -1
        For x = 0 To UBound(FileSystems)
            If FileSystems(x).FileTbl = FileFormat Then
                fs = x
                Exit For
            End If
        Next x
        If fs = -1 Then
            DebugLog("DataReadFileTable", "No such file system as " & CStr(FileFormat), "Check your BeMediaInfoFormats.ini and make sure it isn't corrupt. Alternatively, check to see if you are using an outdated one.")
        End If

        'game search
        gs = -1
        For x = 0 To UBound(Formats)
            If Formats(x).Format = GameFormat Then
                gs = x
                Exit For
            End If
        Next
        If gs = -1 Then
            DebugLog("DataReadFileTable", "No such game database as " & CStr(GameFormat), "Check your BeMediaInfoFormats.ini and make sure it isn't corrupt. Alternatively, check to see if you are using an outdated one.")
        End If

        DataSongSystem = gs
        DataFileSystem = fs
        DataConfig = ConfigToUse

        If fs = -1 Or gs = -1 Then
            DataReadFileTable = True
            Exit Function
        End If

        'chart order setup
        If Formats(gs).KeyS <> "" Then
            a = Formats(gs).KeyS & ","
            l = 0
            Do While a <> ""
                ReDim Preserve ChartAlloc(0 To l)
                x = InStr(a, ",") - 1
                ChartAlloc(l) = Val(Left(a, x))
                a = Mid(a, x + 2)
                l = l + 1
            Loop
            ReDim ChartName((l \ 2) - 1)
            ReDim ChartName5((l \ 2) - 1)
            ReDim DiffIndex((l \ 2) - 1)
            If Formats(gs).DiffS <> "" Then
                a = Formats(gs).DiffS & ","
                l = 0
                Do While a <> ""
                    x = InStr(a, ",") - 1
                    ChartName(l) = (Left(a, x))
                    a = Mid(a, x + 2)
                    l = l + 1
                Loop
            End If
            If Formats(gs).DiffS5 <> "" Then
                a = Formats(gs).DiffS5 & ","
                l = 0
                Do While a <> ""
                    x = InStr(a, ",") - 1
                    ChartName5(l) = (Left(a, x))
                    a = Mid(a, x + 2)
                    l = l + 1
                Loop
            End If
            If Formats(gs).DiffO >= 0 Then
                If Formats(gs).DiffL <> "" Then
                    a = Formats(gs).DiffL & ","
                    l = 0
                    Do While a <> ""
                        x = InStr(a, ",") - 1
                        DiffIndex(l) = Val(Left(a, x))
                        'If DiffIndex(l) = 0 Then DiffIndex(l) = 1
                        a = Mid(a, x + 2)
                        l = l + 1
                    Loop
                End If
            End If
            bChartInfo = True
        End If

        'begin file access
        FileOpen(f, FileName, OpenMode.Binary, OpenAccess.Read, OpenShare.LockWrite)

        'if titles are available, get them here
        If Config(ConfigToUse).Titles >= 0 Then
            o = Config(ConfigToUse).Titles + 1
            Do
                FileGet(f, b, o)
                o += 1
                If b = 0 And d1 = 0 Then
                    Exit Do
                ElseIf b = 0 Then
                    SongNameListCount += 1
                    ReDim Preserve SongNameList(SongNameListCount)
                Else
                    SongNameList(SongNameListCount) &= Chr(b)
                End If
                d1 = b
            Loop
        End If

        'file table
        o = Config(ConfigToUse).TableOffset + 1
        Select Case FileSystems(fs).Count
            Case 1
                FileGet(f, d1, o)
                ind = d1
            Case 2
                FileGet(f, d2, o)
                ind = d2
            Case 4
                FileGet(f, d4, o)
                ind = d4
        End Select
        o += FileSystems(fs).Count
        Do
            If (FileSystems(fs).Count > 0) Then
                If ind > 0 Then
                    ind -= 1
                Else
                    Exit Do
                End If
            End If
            FileGet(f, offs, o + FileSystems(fs).Offset)
            FileGet(f, leng, o + FileSystems(fs).Length)
            If leng = 0 Then
                Exit Do
            End If
            If (CLng(offs) * FileSystems(fs).OffMult) < Offset Then
                Exit Do
            End If
            LastOffset = Offset
            Offset = CLng(offs) * CLng(FileSystems(fs).OffMult)
            Length = CLng(leng) * CLng(FileSystems(fs).LenMult)
            If FileSystems(fs).BackChk <> 0 Then
                If Offset < LastOffset Then
                    Exit Do
                End If
            End If
            If Length <= 0 Then
                Exit Do
            End If
            If Offset >= &H100000000 Or Offset < 0 Then
                '4gb limit
                Exit Do
            End If
            ReDim Preserve DataFiles(0 To DataFileCount)
            With DataFiles(DataFileCount)
                .Offset = Offset
                .Length = Length
                .IsFromInfoFile = False
                .ForcedType = DataDetectedType.None
                .BGM = 0
                .Keysound = 0
                .Difficulty = 0
                .Movie = 0
                .Prefix = ""
                .Name = ""
                .DecodePath = ""
                .DecodeType = SoundDecodeType.None
            End With
            o = o + FileSystems(fs).Size
            DataFileCount += 1
        Loop

        flog = FreeFile()


        'song table
        If Config(ConfigToUse).SongOffset >= 0 Then
            o = Config(ConfigToUse).SongOffset + 1
            With Formats(gs)
                Do
                    ReDim Preserve Songs(0 To SongsCount)
                    ReDim Songs(SongsCount).BGM(0)
                    ReDim Songs(SongsCount).Charts(0)
                    ReDim Songs(SongsCount).Keysounds(0)
                    Songs(SongsCount).Name = "untitledsong" & CStr(SongsCount + Config(ConfigToUse).Index)

                    bValid = True

                    '*** DATA file identification
                    If (.DatFile >= 0) Then
                        FileGet(f, d2, o + .DatFile)
                        If (d2 <> Config(ConfigToUse).DataID) Then
                            bValid = False
                        End If
                    End If



                    '*** NAME
                    If (.NameO >= 0) Or (.NameP <> 0) Then 'name exists?
                        If .NameT = 0 Then 'name type - 0=string
                            ReDim db(0 To .NameSZ - 1)
                            FileGet(f, db, o + .NameO)
                        ElseIf .NameT = 4 Then '4=pointer
                            ReDim db(0 To 63)
                            FileGet(f, d4, o + .NameL)
                            If (d4 <= 0) Or ((.NameP + d4) < 0) Then
                                Exit Do
                            End If
                            FileGet(f, db, d4 + .NameP + 1)
                        End If
                        If db(0) = 0 Then
                            Exit Do
                        End If
                        SongName = ""
                        For x = 0 To UBound(db)
                            If db(x) <> 0 Then
                                SongName = SongName & Chr(db(x))
                            Else
                                Exit For
                            End If
                        Next
                        'now search our current rip list and adjust if there is a match
                        For x = 0 To SongNameListCount - 2
                            If SongNameList(x) = SongName Then
                                SongName = SongNameList(x + 1)
                                Exit For
                            End If
                        Next
                        If Config(ConfigToUse).RipType = "IIDX9" Then
                            'a quick little hack to adjust the names for BMUS
                            If InStr(SongName, "_") > 0 Then
                                SongName = Mid$(SongName, InStr(SongName, "_") + 1)
                            End If
                        End If
                        SongName = DataTranslateName(Trim(SongName))
                        Songs(SongsCount).Name = SongName
                        'scan the SongDB for additional info (if any)
                        SongDBInfo = ConfigGetSongDBInfo(SongName)
                        s = SongDBInfo.InternalName & "|"
                        Do While Len(s) > 1
                            If UCase(Left(s, InStr(s, "|") - 1)) = UCase(SongName) Then
                                SongName = SongDBInfo.Title
                                Exit Do
                            End If
                            s = Mid(s, InStr(s, "|") + 1)
                        Loop
                        If s = "" Then
                            x = x
                        End If
                    End If

                    If bValid Then

                        '*** DIFFICULTY
                        If (.DiffO >= 0) Then
                            ReDim DiffList(0 To .DiffC - 1)
                            For x = 0 To .DiffC - 1
                                FileGet(f, d1, o + .DiffO + x)
                                DiffList(x) = d1
                                'If d1 = 0 Then
                                '    DiffList(x) = 0
                                'End If
                            Next
                        End If

                        '*** MOVIES
                        ' since videos tend to be reused a lot in Bemani games
                        ' we won't be naming them (though you can just uncomment the
                        ' two lines to enable naming again...)
                        If .MovieC > 0 Then
                            If .MovieT = 2 Then
                                lmov = -1
                                ReDim Songs(SongsCount).Movies(0 To .MovieC - 1)
                                For x = 0 To .MovieC - 1
                                    FileGet(f, d2, o + .MovieO + (x * 2))
                                    If d2 < DataFileCount And d2 > 0 Then
                                        'DataFiles(d2).Name = " "
                                        If .MovieD > 0 Then
                                            d2 += .MovieD
                                            'DataFiles(d2).Name = " "
                                        End If
                                        If d2 > -1 Then
                                            Songs(SongsCount).Movies(x) = d2
                                            lmov = d2
                                        Else
                                            Songs(SongsCount).Movies(x) = lmov
                                        End If
                                    End If
                                Next
                            End If
                        End If

                        '*** KEYSOUND/BGM SETS
                        If .SetC > 0 Then
                            If .SetT = 2 Then
                                ReDim KeysList(0 To Formats(gs).SetC - 1)
                                If bChartInfo Then
                                    ReDim Songs(SongsCount).BGM(0 To UBound(ChartName))
                                    ReDim Songs(SongsCount).Keysounds(0 To UBound(ChartName))
                                Else
                                    ReDim Songs(SongsCount).BGM(0 To (.SetC \ 2) - 1)
                                    ReDim Songs(SongsCount).Keysounds(0 To (.SetC \ 2) - 1)
                                    ReDim ChartAlloc(0 To .SetC - 1)
                                    For x = 0 To .SetC - 4 Step 4
                                        ChartAlloc(x + 0) = x
                                        ChartAlloc(x + 1) = x + 2
                                        ChartAlloc(x + 2) = x + 1
                                        ChartAlloc(x + 3) = x + 3
                                    Next
                                    If (.SetC Mod 4) <> 0 Then
                                        For l = ((.SetC \ 4) * 4) To .SetC - 1
                                            ChartAlloc(l) = l
                                        Next
                                    End If
                                End If
                                For l = 0 To .SetC - 1
                                    FileGet(f, d2, o + .SetO + (l * 2))
                                    KeysList(l) = d2 - Config(ConfigToUse).Index
                                Next
                                lkey = KeysList(ChartAlloc(0))
                                lbgm = KeysList(ChartAlloc(1))
                                If lkey >= 0 AndAlso lbgm >= 0 Then
                                    For l = 0 To UBound(ChartAlloc) - 1 Step 4
                                        If KeysList(ChartAlloc(l)) < 0 Then
                                            KeysList(ChartAlloc(l)) = lkey
                                        End If
                                        If KeysList(ChartAlloc(l + 1)) < 0 Then
                                            KeysList(ChartAlloc(l + 1)) = lbgm
                                        End If
                                        If KeysList(ChartAlloc(l + 2)) < 0 Then
                                            KeysList(ChartAlloc(l + 2)) = KeysList(ChartAlloc(l))
                                        End If
                                        If KeysList(ChartAlloc(l + 3)) < 0 Then
                                            KeysList(ChartAlloc(l + 3)) = KeysList(ChartAlloc(l + 1))
                                        End If
                                    Next
                                    For l = 0 To UBound(ChartAlloc) Step 2
                                        If (ChartAlloc(l) >= 0) Or (ChartAlloc(l + 1) >= 0) Then
                                            d2 = KeysList(ChartAlloc(l))
                                            If (d2 > -1) AndAlso (d2 <= UBound(DataFiles)) AndAlso (DataFiles(d2).Name = "") Then
                                                DataFiles(d2).Name = SongName
                                            End If
                                            Songs(SongsCount).Keysounds(l \ 2) = d2
                                            d2 = KeysList(ChartAlloc(l + 1))
                                            If (d2 > -1) AndAlso (d2 <= UBound(DataFiles)) AndAlso (DataFiles(d2).Name = "") Then
                                                DataFiles(d2).Name = SongName
                                            End If
                                            Songs(SongsCount).BGM(l \ 2) = d2
                                        End If
                                    Next

                                    'now apply prefixes to keysound sets if we have multiples
                                    lkey = 0
                                    For l = 0 To UBound(Songs(SongsCount).Keysounds)
                                        For x = 0 To UBound(Songs(SongsCount).Keysounds)
                                            If Songs(SongsCount).Keysounds(x) <= UBound(DataFiles) Then
                                                If (x <> l) And (Songs(SongsCount).Keysounds(x) <> Songs(SongsCount).Keysounds(l)) And DataFiles(Songs(SongsCount).Keysounds(x)).Prefix = "" Then
                                                    If DataFiles(Songs(SongsCount).Keysounds(l)).Prefix = "" Then
                                                        DataFiles(Songs(SongsCount).Keysounds(l)).Prefix = DataBMEString(lkey)
                                                        lkey += 1
                                                    End If
                                                    DataFiles(Songs(SongsCount).Keysounds(x)).Prefix = DataBMEString(lkey)
                                                    lkey += 1
                                                End If
                                            End If
                                        Next
                                    Next

                                    'and further apply a prefix if using 5/7key in BMUS
                                    If .Key5 > 0 Then
                                        FileGet(f, d2, o + .Key5)
                                        b5Key = (d2 = 0)
                                    End If
                                End If

                            End If
                        End If

                        '*** CHARTS
                        ChartTime = 0
                        If (.ChartAdjust > 0) And (.Timing > 0) Then
                            FileGet(f, d2, o + .ChartAdjust)
                            ChartTime = .Timing + d2
                        End If
                        If .KeyC > 0 Then
                            If Config(ConfigToUse).ReadEXE = True Then
                                'read offsets directly from file
                                ReDim Songs(SongsCount).Charts(0 To .KeyC - 1)
                                For x = 0 To .KeyC - 1
                                    FileGet(f, d4, o + .KeyO + (x * 4))
                                    If d4 > 0 Then
                                        ReDim Preserve DataFiles(0 To DataFileCount)
                                        With DataFiles(DataFileCount)
                                            .Name = SongName
                                            If bChartInfo Then
                                                If b5Key Then
                                                    If ChartName5(x) <> "" Then
                                                        .Suffix = " [" & ChartName5(x) & "]"
                                                    End If
                                                Else
                                                    If ChartName(x) <> "" Then
                                                        .Suffix = " [" & ChartName(x) & "]"
                                                    End If
                                                End If
                                                .Keysound = Songs(SongsCount).Keysounds(x)
                                                .BGM = Songs(SongsCount).BGM(x)
                                                .Artist = SongDBInfo.Artist
                                                .Genre = SongDBInfo.Genre
                                                If Formats(gs).MovieO >= 0 Then
                                                    .Movie = Songs(SongsCount).Movies(0)
                                                End If
                                                .ForceTiming = ChartTime
                                            End If
                                            Songs(SongsCount).Charts(x) = DataFileCount
                                            .DecodePath = DataGetFormattedName(CStr(SongsCount))
                                            .IsFromInfoFile = True
                                            .Length = -1
                                            .Offset = (Formats(gs).KeyP + d4)
                                            .ForcedType = DataDetectedType.IIDXOldChart
                                            'apply difficulty
                                            If Formats(gs).DiffL <> "" Then
                                                .Difficulty = DiffList(DiffIndex(x))
                                            End If
                                        End With
                                        DataFileCount += 1
                                    End If
                                Next
                            Else
                                'read indexes
                                If .KeyT = 4 Then
                                    ReDim Songs(SongsCount).Charts(0 To .KeyC - 1)
                                    For x = 0 To .KeyC - 1
                                        FileGet(f, d4, o + .KeyO + (x * 4))
                                        d4 -= Config(ConfigToUse).Index
                                        If d4 < DataFileCount And d4 >= 0 Then
                                            Songs(SongsCount).Charts(x) = d4
                                            DataFiles(d4).Name = SongName
                                            If bChartInfo Then
                                                DataFiles(d4).Keysound = Songs(SongsCount).Keysounds(x)
                                                DataFiles(d4).BGM = Songs(SongsCount).BGM(x)
                                                DataFiles(d4).Movie = Songs(SongsCount).Movies(0)
                                                DataFiles(d4).Artist = SongDBInfo.Artist
                                                DataFiles(d4).Genre = SongDBInfo.Genre
                                                If b5Key Then
                                                    If ChartName5(x) <> "" Then
                                                        DataFiles(d4).Suffix = " [" & ChartName5(x) & "]"
                                                    End If
                                                    'apply difficulty
                                                    '(a little hack for BMUS since the difficulties are out of order)
                                                    If Formats(gs).DiffL <> "" Then
                                                        DataFiles(d4).Difficulty = DiffList(((x Mod 2) * 4) + (x \ 2))
                                                    End If
                                                Else
                                                    If ChartName(x) <> "" Then
                                                        DataFiles(d4).Suffix = " [" & ChartName(x) & "]"
                                                    End If
                                                    'apply difficulty
                                                    If Formats(gs).DiffL <> "" Then
                                                        DataFiles(d4).Difficulty = DiffList(DiffIndex(x))
                                                    End If
                                                End If
                                            Else
                                                DataFiles(d4).Name = SongName
                                            End If
                                            DataFiles(d4).DecodePath = DataGetFormattedName(CStr(SongsCount))
                                        End If
                                    Next
                                End If
                            End If
                        End If
                        SongsCount += 1
                    End If


                    o = o + .Size
                Loop

            End With
            For x = 0 To UBound(DataFiles)
                If DataFiles(x).Name = "" Then
                    DataFiles(x).Name = "untitled" & CStr(x + Config(ConfigToUse).Index)
                End If
                DataFiles(x).Name = Trim(DataFiles(x).Name)
            Next
        End If
        'end file access
        FileClose(f)
        'LogWrite = IO.File.CreateText("c:\log.txt")
        'For x = 0 To UBound(DataFiles)
        '    If DataFiles(x).Name <> "" Then
        '        LogWrite.WriteLine(CStr(x) & " " & DataFiles(x).Name)
        '    End If
        'Next
        'LogWrite.Close()

    End Function

    Public Sub DataByteSwap(ByRef sDat() As Byte)
        Dim a As Byte
        Dim b As Byte
        Dim x As Integer
        If ((UBound(sDat) + 1) And 1) Then
            ReDim Preserve sDat(UBound(sDat) + 1)
        End If
        For x = 0 To UBound(sDat) Step 2
            a = sDat(x)
            b = sDat(x + 1)
            sDat(x) = b
            sDat(x + 1) = a
        Next
    End Sub

    Public Sub DataDecryptIIDXAC(ByVal xSourceData() As Byte, ByRef xTargetData() As Byte, Optional ByVal xKey As String = "", Optional ByVal xDecodeStyle As String = "", Optional ByVal bHeader As Boolean = True)
        Dim DecTag As String
        Dim x As Integer
        Dim o As Long
        Dim DecBlock() As Byte
        Dim LastDecBlock() As Byte
        Dim DecKey() As Byte
        Dim DataOffset As Integer
        ReDim xTargetData(0)
        If xDecodeStyle = "" Then
            If bHeader Then
                DecTag = Chr(xSourceData(0)) & Chr(xSourceData(1)) & Chr(xSourceData(2)) & Chr(xSourceData(3))
                For x = 0 To UBound(DecKeys)
                    With DecKeys(x)
                        If .Signature = DecTag Then
                            xDecodeStyle = .DecodeType
                            xKey = .Block
                            Exit For
                        End If
                    End With
                Next
            End If
        End If
        If xKey <> "" Then
            DecKey = ConfigConvertKeyToBytes(xKey)
        Else
            Exit Sub
        End If
        Select Case xDecodeStyle
            Case "9"
                If bHeader Then
                    DataOffset = 8
                Else
                    DataOffset = 0
                End If
                ReDim DecBlock(0 To 7)
                ReDim LastDecBlock(0 To 7)
                ReDim xTargetData(UBound(xSourceData) - 8)
                For o = DataOffset To UBound(xSourceData) Step 8
                    Array.Copy(xSourceData, o, DecBlock, 0, 8)
                    DataDecryptIIDXACNormal(DecBlock, DecKey)
                    For x = 0 To 7
                        DecBlock(x) = DecBlock(x) Xor LastDecBlock(x)
                    Next
                    Array.Copy(DecBlock, 0, LastDecBlock, 0, 8)
                    Array.Copy(DecBlock, 0, xTargetData, o - DataOffset, 8)
                Next
            Case "12"
        End Select
    End Sub

    Private Sub DataDecryptIIDXACCommon(ByRef xBlock As Byte())
        Dim a, b, c, d, e, f, g, h, i As Integer
        a = (xBlock(0) * 63) And 255
        b = (xBlock(3) + a) And 255
        c = (xBlock(1) * 17) And 255
        d = (xBlock(2) + c) And 255
        e = (d + b) And 255
        f = (xBlock(3) * e) And 255
        g = (f + b + 51) And 255
        h = b Xor d
        i = g Xor e
        xBlock(4) = xBlock(4) Xor h
        xBlock(5) = xBlock(5) Xor d
        xBlock(6) = xBlock(6) Xor i
        xBlock(7) = xBlock(7) Xor g
    End Sub

    Private Sub DataDecryptIIDXACNormal(ByRef xBlock() As Byte, ByRef xKey() As Byte)
        Dim i As Integer
        Dim t As Byte
        For i = 0 To 7
            xBlock(i) = xBlock(i) Xor xKey(i)
        Next
        DataDecryptIIDXACCommon(xBlock)
        For i = 0 To 3
            t = xBlock(i)
            xBlock(i) = xBlock(i + 4)
            xBlock(i + 4) = t
        Next
        For i = 0 To 7
            xBlock(i) = xBlock(i) Xor xKey(i + 8)
        Next
        DataDecryptIIDXACCommon(xBlock)
        For i = 0 To 7
            xBlock(i) = xBlock(i) Xor xKey(i + 16)
        Next
    End Sub

    Public Sub DataDumpArray(ByVal sFileName As String, ByVal xArray() As Byte)
        My.Computer.FileSystem.WriteAllBytes(sFileName, xArray, False)
    End Sub

    Public Sub DataDumpArrayAppend(ByVal sFileName As String, ByVal xArray() As Byte)
        My.Computer.FileSystem.WriteAllBytes(sFileName, xArray, True)
    End Sub

    Public Sub DataDumpArrayXOR(ByVal sFileName As String, ByVal xArray1() As Byte, ByVal xArray2() As Byte)
        Dim x As Integer
        Dim xArraySave(0 To Math.Min(UBound(xArray1), UBound(xArray2))) As Byte
        For x = 0 To UBound(xArraySave)
            xArraySave(x) = xArray1(x) Xor xArray2(x)
        Next
        My.Computer.FileSystem.WriteAllBytes(sFileName, xArraySave, False)
    End Sub

    Public Sub DataExtractRaw(ByVal FileNumber As Integer, ByVal TargetFile As String, ByVal Offset As Long, ByVal Length As Integer)
        If Length <= 0 Then
            Exit Sub
        End If
        Dim b(Length - 1) As Byte
        FileGet(FileNumber, b, Offset + 1, False)
        FileCreateFolder(Strings.Left(TargetFile, InStrRev(TargetFile, "\")))
        My.Computer.FileSystem.WriteAllBytes(TargetFile, b, False)
    End Sub

    Public Sub DataExtractRaw(ByVal FileStream As IO.Stream, ByVal TargetFile As String, ByVal Offset As Long, ByVal Length As Integer)
        Dim OldPos As Long = FileStream.Position
        If Length <= 0 Then
            Exit Sub
        End If
        Dim b(Length - 1) As Byte
        'FileGet(FileNumber, b, Offset + 1, False)
        FileStream.Position = Offset
        FileStream.Read(b, 0, Length)
        FileCreateFolder(Strings.Left(TargetFile, InStrRev(TargetFile, "\")))
        My.Computer.FileSystem.WriteAllBytes(TargetFile, b, False)
        FileStream.Position = OldPos
    End Sub

    Public Function DataExtractFIREBEAT(ByRef CHD As clsCHD, ByVal FileOffset As Long) As Integer
        'sample format: (size=18)
        '0      1      playback channel, FF=auto-assign
        '1      1      unknown (usually 1)
        '2      2      frequency (actual, MSB-LSB)
        '4      1      unknown
        '5      1      pan (01-7F)
        '6      3      offset
        '9      3      length
        '12     1      unknown
        '13     1      unknown
        '14     1      unknown (sampletype?)
        '15     1      flags
        '16     1      volume
        '17     1      unknown
        Dim ThisDataFolder As String
        Dim rbytes() As Byte = {}
        Dim xbytes() As Byte = {}
        Dim SampleInfo() As Byte = {}
        Dim OldSampleInfo() As Byte = {}
        Dim SoundData() As Byte = {}
        Dim ChartData() As Byte = {}
        Dim BGMLeft() As Byte = {}
        Dim BGMRight() As Byte = {}
        Dim BGMLeftS() As Short = {}
        Dim BGMRightS() As Short = {}
        Dim Sample() As Byte = {}
        Dim VolumeL As Double
        Dim VolumeR As Double
        Dim SampleOffset As Integer
        Dim SampleOffset2 As Integer
        Dim SoundDataSize As Integer
        Dim SampleLength As Integer
        Dim SampleLength2 As Integer
        Dim SampleEnd As Integer
        Dim SampleEnd2 As Integer
        Dim Frequency As Integer
        Dim HasChartData As Boolean
        Dim IsValidChart As Boolean
        Dim SoundCheckCount As Integer
        Dim x As Integer
        Dim y As Integer
        CHD.ReadHunkBytes(FileOffset, 4, rbytes)
        DataExtractFIREBEAT = 0

        'these are all invalid for everything, so don't bother
        If rbytes(0) = 0 And rbytes(1) = 0 And rbytes(2) = 0 And rbytes(3) = 0 Then
            Exit Function
        End If
        If rbytes(0) = &H88 And rbytes(1) = &H88 And rbytes(2) = &H88 And rbytes(3) = &H88 Then
            Exit Function
        End If
        If rbytes(0) = &HA And rbytes(1) = &HA And rbytes(2) = &HA And rbytes(3) = &HA Then
            Exit Function
        End If
        If rbytes(0) = &H4F And rbytes(1) = &H4F And rbytes(2) = &H4F And rbytes(3) = &H4F Then
            Exit Function
        End If

        ThisDataFolder = sTargetFolder & DataPadString(Hex(FileOffset), 10, "0") & "\"

        CHD.ReadHunkBytes(FileOffset, 18, rbytes)
        DataByteSwap(rbytes)

        ReDim OldSampleInfo(&H2000 - 1)
        If rbytes(1) > 0 And (rbytes(2) > 0 Or rbytes(3) > 0) And rbytes(4) > 0 And (rbytes(5) > 0 And rbytes(5) < &H80) And rbytes(14) = 7 Then
            CHD.ReadHunkBytes(FileOffset, &H2000, SampleInfo)
            DataByteSwap(SampleInfo)
            'determine sample set size
            For x = 0 To (256 * 18) - 1 Step 18
                If SampleInfo(x + 14) = 7 Then
                    SampleEnd = DataMakeInt32(SampleInfo(x + 11), SampleInfo(x + 10), SampleInfo(x + 9)) * 2
                    If SampleEnd > SoundDataSize Then
                        SoundDataSize = SampleEnd
                    End If
                Else
                    Exit For
                End If
            Next
            If SoundDataSize > 0 Then
                FileCreateFolder(ThisDataFolder)
                'find and extract charts
                For x = 0 To 31
                    If (x = 6) And (Not HasChartData) Then
                        Exit For
                    End If
                    IsValidChart = False
                    CHD.ReadHunkBytes(FileOffset + &H2000 + (x * &H4000), &H4000, ChartData)
                    If ChartData(0) = 0 And ChartData(1) = 0 And ChartData(2) = 0 And ChartData(3) > 0 And ChartData(3) <= 250 And ChartData(4) = 0 Then
                        For y = 0 To &H3FFC Step 4
                            If ChartData(y) = &HFF And ChartData(y + 1) = &H7F And ChartData(y + 2) = 0 And ChartData(y + 3) = 0 Then
                                ReDim Preserve ChartData(y + 3)
                                HasChartData = True
                                IsValidChart = True
                                Exit For
                            End If
                        Next
                        If IsValidChart Then
                            If ThisJob.RipCharts Then
                                If ThisJob.ConvertChart Then
                                    ChartLoadMemory(ChartData, ChartTypes.IIDXCS2)
                                    ChartSaveBMS(ThisDataFolder & "chart" & CStr(x) & ".bme", ChartTypes.IIDXCS2, 192)
                                Else
                                    DataDumpArray(ThisDataFolder & "chart" & CStr(x) & ".cs2", ChartData)
                                End If
                            End If
                        End If
                    End If
                Next
                'the presence of a chart determines where the sound data is located
                If HasChartData Then
                    If ThisJob.RipKeysounds Then
                        CHD.ReadHunkBytes(FileOffset + &H100000, SoundDataSize + 1, SoundData)
                    End If
                    DataExtractFIREBEAT = SoundDataSize + &H100000
                    SoundCheckCount = 1
                Else
                    If ThisJob.RipKeysounds Then
                        CHD.ReadHunkBytes(FileOffset + &H20000, SoundDataSize + 1, SoundData)
                    End If
                    DataExtractFIREBEAT = SoundDataSize + &H20000
                    SoundCheckCount = 1
                End If

                For y = 0 To SoundCheckCount - 1
                    CHD.ReadHunkBytes(FileOffset + (y * &H18000), &H2000, SampleInfo)
                    DataByteSwap(SampleInfo)
                    DataDumpArray(ThisDataFolder & "!samples" & CStr(y) & ".info", SampleInfo)
                    If ThisJob.RipKeysounds Then
                        For x = 0 To (256 * 18) - 1 Step 18
                            If SampleInfo(x + 14) = 7 Then
                                VolumeL = PanTableRight(((SampleInfo(x + 5) - 1) / &H7E) * &H100)
                                VolumeR = PanTableLeft(((SampleInfo(x + 5) - 1) / &H7E) * &H100)
                                VolumeL *= FirebeatVolumeTable(SampleInfo(x + 4))
                                VolumeR *= FirebeatVolumeTable(SampleInfo(x + 4))
                                SampleOffset = DataMakeInt32(SampleInfo(x + 8), SampleInfo(x + 7), SampleInfo(x + 6)) * 2
                                SampleEnd = DataMakeInt32(SampleInfo(x + 11), SampleInfo(x + 10), SampleInfo(x + 9)) * 2
                                SampleLength = ((SampleEnd - SampleOffset) + 1)
                                SampleOffset2 = DataMakeInt32(SampleInfo(x + 26), SampleInfo(x + 25), SampleInfo(x + 24)) * 2
                                SampleEnd2 = DataMakeInt32(SampleInfo(x + 29), SampleInfo(x + 28), SampleInfo(x + 27)) * 2
                                SampleLength2 = (SampleEnd2 - SampleOffset2) + 1
                                Frequency = DataMakeInt32(SampleInfo(x + 3), SampleInfo(x + 2))
                                If (SampleOffset <> SampleOffset2) And (SampleInfo(x + 32) = 7) And ((SampleInfo(x + 15) And &H80) = 0) And (SampleLength2 = SampleLength) And ((SampleInfo(x + 5) <= 2) And (SampleInfo(x + 23) >= &H7E)) Then
                                    'stereo BGM
                                    VolumeR = PanTableLeft(((SampleInfo(x + 23) - 1) / &H7E) * &H100)
                                    VolumeR *= FirebeatVolumeTable(SampleInfo(x + 22))
                                    If SampleLength <> SampleLength2 Then
                                        SampleLength = Math.Max(SampleLength, SampleLength2)
                                    End If
                                    ReDim BGMLeft(SampleLength - 1)
                                    ReDim BGMRight(SampleLength - 1)
                                    Array.ConstrainedCopy(SoundData, SampleOffset, BGMLeft, 0, SampleLength)
                                    Array.ConstrainedCopy(SoundData, SampleOffset2, BGMRight, 0, SampleLength)
                                    SoundUpsample(BGMLeft, BGMLeftS, True, 16, 1, Frequency, VolumeL, VolumeR)
                                    SoundUpsample(BGMRight, BGMRightS, True, 16, 1, Frequency, VolumeL, VolumeR)
                                    SoundRemoveSilence(BGMLeftS)
                                    SoundRemoveSilence(BGMRightS)
                                    SoundCombineSave(BGMLeftS, BGMRightS, ThisDataFolder & DataBMEString((x \ 18) + 1) & ".wav", Frequency)
                                    x += 18
                                ElseIf (SampleInfo(x + 15) And &H80) = 0 Then
                                    'mono sample
                                    ReDim Sample(SampleLength - 1)
                                    Array.ConstrainedCopy(SoundData, SampleOffset, Sample, 0, SampleLength)
                                    SoundUpsampleSave(Sample, ThisDataFolder & DataBMEString((x \ 18) + 1) & ".wav", True, 16, 1, Frequency, VolumeL, VolumeR)
                                Else
                                    'stereo sample
                                    ReDim Sample(SampleLength - 1)
                                    Array.ConstrainedCopy(SoundData, SampleOffset, Sample, 0, SampleLength)
                                    SoundUpsampleSave(Sample, ThisDataFolder & DataBMEString((x \ 18) + 1) & ".wav", True, 16, 2, Frequency, VolumeL, VolumeR)
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If
                Next
            End If
        Else
            x = x
        End If
        BGMLeftS = Nothing
        BGMRightS = Nothing
        SoundData = Nothing
        Sample = Nothing
        ChartData = Nothing
        BGMLeft = Nothing
        BGMRight = Nothing
    End Function

    Public Sub DataExtractDJMAIN(ByRef CHD As clsCHD, ByVal FileOffset As Long)
        'sample format: (size=11)
        '0      1      playback channel (two sounds can not use the same channel)
        '1      2      frequency - to get this for WAV you can use freq=((x/60216)*44100)
        '3      1      reverb volume (not typically used in beatmania)
        '4      1      volume
        '5      1      pan 1-F (lower 4 bits ONLY)
        '6      3      offset
        '9      1      sample type flag:
        '                0 - signed 8bit (ends 80 80 80 80 80 80 80 80)
        '                4 - signed 16bit (ends 00 80 00 80 00 80 00 80 00 80 00 80 00 80 00 80)
        '                8 - delta 4bit (ends 88 88 88 88)
        '10     1      flags:
        '                1 - this sample loops: (we should never encounter these in beatmania songs)
        '                    16bit marker = 00 80
        '                     8bit marker = 80
        '                     4bit marker = 88
        '              128 - unknown but it shows up frequently
        Dim SampleInfoSize As Integer = 11
        Dim rbytes() As Byte = {}
        Dim xbytes() As Byte = {}
        Dim sampleinfo() As Byte = {}
        Dim BaseOffset As Long = FileOffset
        Dim ChartInfo() As Byte
        Dim MaxRead As Integer
        'Dim TempOffset As Integer
        Dim ThisDataFolder As String
        Dim HasSampleData As Boolean = False
        Dim HasChartData As Boolean = False
        Dim ThisKeysoundSet As Integer = 0
        Dim Prefix As String
        Dim bExtracted As Boolean
        Dim SampleOffset As Integer
        Dim ChartName As String
        Dim SearchOffset As Integer
        Dim KeysoundSets As Integer
        Dim AudioData() As Byte = {} '16mb max
        Dim AudioConvert() As Byte
        Dim LastSampleInfo() As Byte = {}
        Dim SkipAmount As Integer
        Dim BGMLeft() As Short = {}
        Dim BGMRight() As Short = {}
        Dim ChartOffset As Integer
        Dim ConvertedSample() As Short = {}
        Dim CombineBGMs As Boolean
        Dim ThisSampleNumber As Integer
        Dim DontConvertYet As Boolean
        Dim HaveLeftBGM As Boolean = False
        Dim HaveRightBGM As Boolean = False
        Dim HaveCorrectSample As Boolean = False
        Dim NoMoreBGMs As Boolean
        Dim RipThisData As Boolean
        Dim DifferentKeyInfo As Boolean
        Dim SkipThisSample As Boolean
        Dim BGMFile As String = ""
        Dim BGMNumber As Integer = 0
        Dim AudioRipped As Boolean = False
        Dim SoundOffset As Integer
        Dim MainOffset As Integer = 0
        Dim MaxChartSize As Integer = &H2000
        Dim GameString As String = ""
        Dim HighestUnusedPart As Integer
        Dim HighestData As Integer
        Dim FirstSoundTableOffset As Integer = -1
        Dim HighestSampleOffset As Integer = 0
        Dim HighestTypeOffset As Integer = 0
        Dim LastSampleType As Byte
        Dim ThisSampleType As Byte
        Dim MultipleSoundSets As Boolean
        Dim ThisSection As Integer = FileOffset \ &H1000000
        Dim ChartMetaData As String = ""
        Dim ChartDBInfo As ConfigSongDBType = SongDB(0)
        Dim CheckDBInfo As ConfigSongDBType = SongDB(0)
        Dim bFoundTagInfo As Boolean
        Dim ChartReportedNoteCount As Integer
        Dim ChartReportedBPM As Integer
        Dim ChartNoteCounts() As Integer
        Dim bIsDoubleSet As Boolean


        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim a As Integer
        Dim volLeft As Double
        Dim volRight As Double
        Dim Freq As Integer

        SoundRipInfo.DoRip = False
        Prefix = ""

        CHD.ReadHunkBytes(FileOffset, 4, rbytes)

        'these are all invalid for everything, so don't bother
        If rbytes(0) = 0 And rbytes(1) = 0 And rbytes(2) = 0 And rbytes(3) = 0 Then
            Exit Sub
        End If
        If rbytes(0) = &H88 And rbytes(1) = &H88 And rbytes(2) = &H88 And rbytes(3) = &H88 Then
            Exit Sub
        End If
        If rbytes(0) = &HA And rbytes(1) = &HA And rbytes(2) = &HA And rbytes(3) = &HA Then
            Exit Sub
        End If
        If rbytes(0) = &H4F And rbytes(1) = &H4F And rbytes(2) = &H4F And rbytes(3) = &H4F Then
            Exit Sub
        End If

        'CHD.ReadHunkBytes(FileOffset, &H1000000, xbytes)
        'DataDumpArray(sTargetFolder & DataPadString(Hex(FileOffset), 10, "0") & ".dat", xbytes)
        'If True Then
        '    Exit Sub
        'End If


        sDecoderInfo = "Scanning for data"

        ReDim ChartInfo(255)
        Dim ChartLocations(255) As Boolean
        Dim KeyLocations(255) As Boolean

        'DJMAIN tends to move offsets around
        'FIREBEAT tends to stay consistent

        'CHD.ReadHunkBytes(FileOffset, &H20000, xbytes)
        'DataDumpArray(sTargetFolder & DataPadString(Hex(FileOffset), 10, "0") & ".hdr", xbytes)
        'Exit Sub

        SkipAmount = &H200
        HighestUnusedPart = 0
        y = 0
        Do
            CHD.ReadHunkBytes(FileOffset + y, SkipAmount, xbytes)
            ReDim rbytes(xbytes.Length - 1)
            Array.Copy(xbytes, rbytes, xbytes.Length)
            DataByteSwap(rbytes)
            RipThisData = False
            For x = 0 To SkipAmount - 1
                If rbytes(x) <> 0 And rbytes(x) <> &HA And rbytes(x) <> &H4F And rbytes(x) <> &H88 Then
                    RipThisData = True
                    Exit For
                End If
            Next
            If RipThisData Then 'Not ((rbytes(0) = rbytes(1) And rbytes(1) = rbytes(2) And rbytes(2) = rbytes(3)) And (rbytes(0) = 0 Or rbytes(0) = &H88 Or rbytes(0) = &HA Or rbytes(0) = &H4F)) Then
                If xbytes(0) = 0 And xbytes(1) = 0 And (xbytes(2) = 0 Or xbytes(2) = &H10 Or xbytes(2) = 2) And (xbytes(3) > 0 And xbytes(3) <= 250) Then
                    'possibly found a chart here
                    CHD.ReadHunkBytes(FileOffset + y, MaxChartSize, rbytes)
                    For x = 0 To MaxChartSize - 4 Step 4
                        If rbytes(x + 0) = &HFF And rbytes(x + 1) = &H7F And rbytes(x + 2) = 0 And rbytes(x + 3) = 0 Then
                            'end of chart
                            HasChartData = True
                            ChartLocations(y \ SkipAmount) = True
                            HighestData = y
                            y += (x \ SkipAmount) * SkipAmount
                            Exit For
                        End If
                    Next
                ElseIf ((rbytes(0) <> 0) Or (rbytes(1) <> 0) Or (rbytes(2) <> 0)) And (rbytes(9) < &H10) And ((rbytes(5) >= &H81 And rbytes(5) <= &H8F)) Then
                    'possibly found a keysound table here
                    If FirstSoundTableOffset < 0 Then
                        FirstSoundTableOffset = y
                    Else
                        MultipleSoundSets = True
                    End If
                    CHD.ReadHunkBytes(FileOffset + y, MaxChartSize, rbytes)
                    DataByteSwap(rbytes)
                    For x = 0 To MaxChartSize - SampleInfoSize Step SampleInfoSize
                        If Not ((rbytes(x + 5) >= &H81 And rbytes(x + 5) <= &H8F) And (rbytes(x + 1) <> 0 Or rbytes(x + 2) <> 0) And (rbytes(x + 9) < &H10)) Then
                            'end of keysound table
                            HasSampleData = True
                            KeyLocations(y \ SkipAmount) = True
                            HighestData = y
                            y += (x \ SkipAmount) * SkipAmount
                            'SoundOffset = y + SkipAmount
                            Exit For
                        End If
                    Next
                Else
                    'can't identify this...
                    x = x
                End If
            Else
                HighestUnusedPart = y
            End If
            y += SkipAmount
        Loop While (y \ SkipAmount) < 256

        'CHD.ReadHunkBytes(FileOffset + SoundOffset, SkipAmount, xbytes)

        'offset checking!
        RipThisData = False
        LastSampleType = 0
        If FirstSoundTableOffset >= 0 Then
            CHD.ReadHunkBytes(FileOffset + FirstSoundTableOffset, &H1000, rbytes)
            DataByteSwap(rbytes)
            'first, find the highest valued sample
            For x = 0 To &H1000 - (SampleInfoSize + 1) Step SampleInfoSize
                If rbytes(x + 5) >= &H81 And rbytes(x + 6) <= &H8F Then
                    SampleOffset = DataMakeInt32(rbytes(x + 6), rbytes(x + 7), rbytes(x + 8))
                    If (SampleOffset > HighestSampleOffset) And SampleOffset < &H1000000 Then
                        CHD.ReadHunkBytes(FileOffset + SampleOffset + &H400, 4, xbytes)
                        If Not (xbytes(0) = &H88 And xbytes(1) = &H88 And xbytes(2) = &H88 And xbytes(3) = &H88) Then
                            HighestSampleOffset = SampleOffset
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            'and the one before that so we can check previous type
            For x = 0 To &H1000 - (SampleInfoSize + 1) Step SampleInfoSize
                If rbytes(x + 5) >= &H81 And rbytes(x + 6) <= &H8F Then
                    SampleOffset = DataMakeInt32(rbytes(x + 6), rbytes(x + 7), rbytes(x + 8))
                    If (SampleOffset < HighestSampleOffset) And (SampleOffset > HighestTypeOffset) Then
                        ThisSampleType = rbytes(x + 9)
                        HighestTypeOffset = SampleOffset
                    End If
                Else
                    Exit For
                End If
            Next
            If HighestSampleOffset > 0 Then
                If (HighestSampleOffset And 1) Then
                    HighestSampleOffset += 1
                End If
                'check a couple "known" sound offsets for a sample-ending just prior to the sample
                For x = 0 To ((&H20000 - HighestData) \ SkipAmount)
                    If (x = 0) And (HighestData <= &H2000) Then
                        SearchOffset = &H2000
                    ElseIf (x = 1) And (HighestData <= &H400) Then
                        SearchOffset = &H400
                    ElseIf x = 2 And (HighestData <= &H20000) Then
                        SearchOffset = &H20000
                    ElseIf x > 2 Then
                        SearchOffset = 0
                        'SearchOffset = HighestData + (x * SkipAmount)
                    Else
                        SearchOffset = 0
                    End If
                    If SearchOffset > 0 Then
                        CHD.ReadHunkBytes((FileOffset + SearchOffset + HighestSampleOffset) - 8, 8, rbytes)
                        DataByteSwap(rbytes)
                        Select Case ThisSampleType
                            Case 0
                                If rbytes(0) = &H80 And rbytes(1) = &H80 And rbytes(2) = &H80 And rbytes(3) = &H80 And rbytes(4) = &H80 And rbytes(5) = &H80 Then
                                    SoundOffset = SearchOffset
                                    RipThisData = True
                                    Exit For
                                End If
                            Case 4
                                If rbytes(0) = 0 And rbytes(1) = &H80 And rbytes(2) = 0 And rbytes(3) = &H80 And rbytes(4) = 0 And rbytes(5) = &H80 Then
                                    SoundOffset = SearchOffset
                                    RipThisData = True
                                    Exit For
                                End If
                            Case 8
                                If rbytes(0) <> &H88 And rbytes(4) = &H88 And rbytes(5) = &H88 And rbytes(6) = &H88 Then
                                    SoundOffset = SearchOffset
                                    RipThisData = True
                                    Exit For
                                End If
                        End Select
                    End If
                Next
            End If
        End If

        If SoundOffset = &H2000 Or SoundOffset = &H20000 Then
            LastDJMainOffset = SoundOffset
        End If

        If Not RipThisData Then
            If (LastDJMainOffset = 0) Then
                SoundOffset = HighestUnusedPart + SkipAmount
            Else
                If (HighestUnusedPart + SkipAmount) = &H400 Then
                    SoundOffset = &H400
                Else
                    SoundOffset = LastDJMainOffset
                    If (HighestUnusedPart + SkipAmount) <> LastDJMainOffset Then
                        x = x
                    End If
                End If
            End If
        End If

        ReDim ChartNoteCounts((2 * SoundOffset) \ SkipAmount)
        'song DB checking
        x = 0
        For y = 0 To (SoundOffset - 1) Step SkipAmount
            For a = 0 To &HF02000 Step &HF02000
                CHD.ReadHunkBytes(FileOffset + y + a, MaxChartSize, rbytes)
                ChartReportedNoteCount = 0
                If rbytes(0) = 0 And rbytes(1) = 0 And rbytes(2) = 0 And rbytes(3) <> 0 Then
                    'determine note count
                    For z = 0 To MaxChartSize - 4 Step 4
                        If (rbytes(z) = 0) And (rbytes(z + 1) = 0) And ((rbytes(z + 2) = 0) Or (rbytes(z + 2) = 16)) And (rbytes(z + 3) <= 250) And (rbytes(z + 3) > 0) Then
                            ChartReportedNoteCount += rbytes(z + 3)
                            If rbytes(z + 2) = 16 Then
                                bIsDoubleSet = True
                            End If
                            'this is a ridiculous hack, sometimes audio data looks like charts!
                            If (rbytes(z + 3) < 250) And (rbytes(z + 6) = 0) Then
                                ChartReportedNoteCount = 0
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If (ChartReportedNoteCount > 0) Then
                        ChartNoteCounts(x) = ChartReportedNoteCount
                        x += 1
                        If ChartReportedBPM = 0 Then
                            'determine bpm
                            For z = 0 To MaxChartSize - 4 Step 4
                                If (rbytes(z) = 0) And (rbytes(z + 1) = 0) And ((rbytes(z + 2) And 15) = 2) Then
                                    ChartReportedBPM = rbytes(z + 3)
                                    Exit For
                                End If
                                If (rbytes(z) = &HFF) And (rbytes(z + 1) = &H7F) Then
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            Next
        Next
        'check the database
        CheckDBInfo = ConfigGetSongDBInfo(ChartNoteCounts, ChartReportedBPM)
        If CheckDBInfo.Title <> "" Then
            ChartDBInfo = CheckDBInfo
            bFoundTagInfo = True
            ThisDataFolder = sTargetFolder & DataPadString(Hex(FileOffset), 10, "0") & " " & DataGetFormattedName(CheckDBInfo.Title) & "\"
        Else
            ThisDataFolder = sTargetFolder & DataPadString(Hex(FileOffset), 10, "0") & "\"
        End If


        FileCreateFolder(ThisDataFolder)


        HasChartData = False
        HasSampleData = False

        For y = 0 To SoundOffset - 1 Step SkipAmount
            CHD.ReadHunkBytes(FileOffset + y, MaxChartSize, rbytes)

            'sample list
            If KeyLocations(y \ SkipAmount) Then '(ChartLocations(y \ SkipAmount) = False) And ((rbytes(0) <> 0) Or (rbytes(1) <> 0) Or (rbytes(3) <> 0)) And (rbytes(8) < &H10) And ((rbytes(4) >= &H81 And rbytes(4) <= &H8F)) Then
                sampleinfo = rbytes
                RipThisData = False
                DataByteSwap(sampleinfo)
                For x = 0 To MaxChartSize - (SampleInfoSize + 1) Step SampleInfoSize
                    'If (sampleinfo(x + 10) = &HA) And (sampleinfo(x) = &HA) And (sampleinfo(x + 6) = &HA And sampleinfo(x + 7) = &HA And sampleinfo(x + 8) = &HA) Then
                    If Not ((sampleinfo(x + 5) >= &H81 And sampleinfo(x + 5) <= &H8F)) Then
                        If x > 0 Then
                            ReDim Preserve sampleinfo(x - 1)
                            RipThisData = True
                        Else
                            ReDim sampleinfo(0)
                        End If
                        Exit For
                    End If
                Next
                If HasSampleData Then
                    DifferentKeyInfo = False
                    If UBound(LastSampleInfo) = UBound(sampleinfo) Then
                        For x = 0 To UBound(sampleinfo) Step 11
                            If LastSampleInfo(x + 6) <> sampleinfo(x + 6) Then
                                DifferentKeyInfo = True
                                Exit For
                            End If
                        Next
                    Else
                        DifferentKeyInfo = True
                    End If
                    If (Not DifferentKeyInfo) Then
                        RipThisData = False
                    Else
                        Prefix = Chr(97 + ThisKeysoundSet)
                        ThisKeysoundSet += 1
                    End If
                End If

                ReDim LastSampleInfo(UBound(sampleinfo))
                sampleinfo.CopyTo(LastSampleInfo, 0)

                'debug...
                'DataDumpArray(ThisDataFolder & "sounds" & DataPadString(Hex(y \ SkipAmount), 2, "0") & ".info", sampleinfo)
                'RipThisData = False

                If RipThisData Then
                    HasSampleData = True
                    KeysoundSets += 1
                    If Not (HaveLeftBGM And HaveRightBGM) Then
                        HaveCorrectSample = True
                        For x = 0 To UBound(sampleinfo) - (SampleInfoSize - 1) Step SampleInfoSize
                            SampleOffset = DataMakeInt32(sampleinfo(x + 6), sampleinfo(x + 7), sampleinfo(x + 8))
                            If (SampleOffset = 0) Then
                                HaveLeftBGM = True
                            ElseIf (SampleOffset = &H680000) Then
                                HaveRightBGM = True
                            ElseIf (SampleOffset > &H680000 And SampleOffset < &HD00000) Or (SampleOffset > 0 And SampleOffset < &H680000) Then
                                'don't do BGM combining with this one, samples are at impossible offsets
                                HaveLeftBGM = False
                                HaveRightBGM = False
                                HaveCorrectSample = False
                                Exit For
                            End If
                        Next
                        CombineBGMs = (HaveLeftBGM And HaveRightBGM And HaveCorrectSample)
                        HaveLeftBGM = False
                        HaveRightBGM = False
                    End If
                    If (AudioRipped = False) And ThisJob.RipKeysounds Then
                        sDecoderInfo = "Extracting Keysound Sectors"
                        If bFoundTagInfo Then
                            sDecoderInfo &= " [" & ChartDBInfo.Title & "]"
                        End If
                        ReDim AudioData(0 To (&H1000000 - SoundOffset) - 1)
                        CHD.ReadHunkBytes(BaseOffset + SoundOffset, UBound(AudioData) + 1, AudioData)
                        DataByteSwap(AudioData)
                        AudioRipped = True
                        MaxRead = UBound(AudioData)
                    End If
                    sDecoderInfo = "Extracting Keysound Set +" & Hex(y)
                    If bFoundTagInfo Then
                        sDecoderInfo &= " [" & ChartDBInfo.Title & "]"
                    End If
                    For x = 0 To UBound(sampleinfo) - (SampleInfoSize - 1) Step SampleInfoSize
                        SkipThisSample = False
                        Freq = Int((DataMakeInt32(sampleinfo(x + 1), sampleinfo(x + 2)) / 60216) * 44100)
                        If Freq <= 0 Then Freq = 44100
                        bExtracted = False
                        SampleOffset = DataMakeInt32(sampleinfo(x + 6), sampleinfo(x + 7), sampleinfo(x + 8))
                        If Not (HaveLeftBGM And HaveRightBGM) Then
                            DontConvertYet = ((CombineBGMs = True) And ((SampleOffset = 0) Or (SampleOffset = &H680000)))
                        Else
                            DontConvertYet = False
                        End If
                        If NoMoreBGMs Then
                            If (SampleOffset = 0) Or (SampleOffset = &H680000) Then
                                SkipThisSample = True
                            End If
                        End If
                        SoundFileProgress = Int((x / UBound(sampleinfo)) * 100)
                        volLeft = PanTableLeft((((sampleinfo(x + 5) And 15) - 1) / &HE) * &H100)
                        volRight = PanTableRight((((sampleinfo(x + 5) And 15) - 1) / &HE) * &H100)
                        volLeft *= DJMainVolumeTable(sampleinfo(x + 4))
                        volRight *= DJMainVolumeTable(sampleinfo(x + 4))
                        If (Not SkipThisSample) Then
                            ThisSampleNumber = (x \ 11) + 2
                            If ThisJob.RipKeysounds Then
                                Select Case sampleinfo(x + 9)
                                    Case 0 '8bit
                                        For z = SampleOffset + 1 To MaxRead - 8
                                            If (AudioData(z) = &H80) And (AudioData(z + 1) = &H80) And (AudioData(z + 2) = &H80) And (AudioData(z + 3) = &H80) And (AudioData(z + 4) = &H80) And (AudioData(z + 5) = &H80) And (AudioData(z + 6) = &H80) And (AudioData(z + 7) = &H80) Then
                                                If ThisJob.RipKeysounds Then
                                                    ReDim AudioConvert((z - SampleOffset) - 1)
                                                    Array.ConstrainedCopy(AudioData, SampleOffset, AudioConvert, 0, (z - SampleOffset - 1))
                                                    SoundUpsample(AudioConvert, ConvertedSample, False, 8, 1, Freq, volLeft, volRight)
                                                End If
                                                bExtracted = True
                                                Exit For
                                            End If
                                        Next
                                    Case 4 '16bit
                                        For z = SampleOffset + 2 To MaxRead - 16
                                            If AudioData(z) = &H0 And AudioData(z + 1) = &H80 And AudioData(z + 2) = &H0 And AudioData(z + 3) = &H80 And AudioData(z + 4) = &H0 And AudioData(z + 5) = &H80 And AudioData(z + 6) = &H0 And AudioData(z + 7) = &H80 Then
                                                If AudioData(z + 8) = &H0 And AudioData(z + 9) = &H80 And AudioData(z + 10) = &H0 And AudioData(z + 11) = &H80 And AudioData(z + 12) = &H0 And AudioData(z + 13) = &H80 And AudioData(z + 14) = &H0 And AudioData(z + 15) = &H80 Then
                                                    If ThisJob.RipKeysounds Then
                                                        ReDim AudioConvert((z - SampleOffset) - 1)
                                                        Array.ConstrainedCopy(AudioData, SampleOffset, AudioConvert, 0, (z - SampleOffset - 1))
                                                        SoundUpsample(AudioConvert, ConvertedSample, False, 16, 1, Freq, volLeft, volRight)
                                                    End If
                                                    bExtracted = True
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    Case 8 '4bit
                                        For z = SampleOffset + 1 To MaxRead - 4
                                            If AudioData(z) = &H88 And AudioData(z + 1) = &H88 And AudioData(z + 2) = &H88 And AudioData(z + 3) = &H88 Then
                                                If ThisJob.RipKeysounds Then
                                                    ReDim AudioConvert((z - SampleOffset) - 1)
                                                    Array.ConstrainedCopy(AudioData, SampleOffset, AudioConvert, 0, (z - SampleOffset - 1))
                                                    SoundUpsample(AudioConvert, ConvertedSample, False, 4, 1, Freq, volLeft, volRight)
                                                End If
                                                bExtracted = True
                                                Exit For
                                            End If
                                        Next
                                End Select
                            End If
                            If bExtracted Or (Not ThisJob.RipKeysounds) Then
                                If Not DontConvertYet Then
                                    If ThisJob.RipKeysounds Then
                                        SoundRemoveSilence(ConvertedSample)
                                        SoundSave(ConvertedSample, ThisDataFolder & Prefix & DataBMEString(ThisSampleNumber) & ".wav", Freq)
                                    End If
                                Else
                                    If ((sampleinfo(x + 5) And 15) >= 8) And (HaveLeftBGM = False) Then
                                        BGMLeft = ConvertedSample
                                        HaveLeftBGM = True
                                    ElseIf ((sampleinfo(x + 5) And 15) <= 8) And (HaveRightBGM = False) Then
                                        BGMRight = ConvertedSample
                                        HaveRightBGM = True
                                    End If
                                    If HaveLeftBGM And HaveRightBGM Then
                                        If Not NoMoreBGMs Then
                                            If ThisJob.RipKeysounds Then
                                                SoundRemoveSilence(BGMLeft)
                                                SoundRemoveSilence(BGMRight)
                                                SoundCombineSave(BGMLeft, BGMRight, ThisDataFolder & "@BGM.wav", Freq)
                                            End If
                                            BGMFile = "@BGM.wav"
                                            NoMoreBGMs = True
                                        End If
                                        BGMNumber = ThisSampleNumber
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Next

        'now extract the charts
        For z = 0 To &HF02000 Step &HF02000
            For y = 0 To SoundOffset - 1 Step SkipAmount
                'chart
                If ((z > 0) Or (ChartLocations(y \ SkipAmount) And (z = 0))) And ThisJob.RipCharts Then
                    ChartOffset = z + y
                    CHD.ReadHunkBytes(FileOffset + ChartOffset, MaxChartSize, rbytes)
                    If rbytes(0) = 0 And rbytes(1) = 0 Then
                        For x = 0 To MaxChartSize - 5 Step 4
                            If rbytes(x) = &HFF And rbytes(x + 1) = &H7F And rbytes(x + 2) = 0 And rbytes(x + 3) = 0 Then
                                ChartInfo = rbytes
                                ReDim Preserve ChartInfo(x + 3)
                                ChartMetaData = ""
                                Select Case ChartOffset
                                    Case &H400, &H800
                                        ChartName = " [" & IIf(bIsDoubleSet, "DP", "SP") & " standard]"
                                        ChartMetaData &= "#DIFFICULTY 3" & vbCrLf
                                    Case &HF02000
                                        ChartName = " [" & IIf(bIsDoubleSet, "DP", "SP") & " basic]"
                                        ChartMetaData &= "#DIFFICULTY 2" & vbCrLf
                                    Case &HF03000
                                        ChartName = " [" & IIf(bIsDoubleSet, "DP", "SP") & " another]"
                                        ChartMetaData &= "#DIFFICULTY 4" & vbCrLf
                                    Case &H2000
                                        ChartName = " [SP standard]"
                                        ChartMetaData &= "#DIFFICULTY 3" & vbCrLf
                                    Case &H6000
                                        ChartName = " [SP basic]"
                                        ChartMetaData &= "#DIFFICULTY 2" & vbCrLf
                                    Case &HA000
                                        ChartName = " [SP another]"
                                        ChartMetaData &= "#DIFFICULTY 4" & vbCrLf
                                    Case &HE000
                                        ChartName = " [DP standard]"
                                        ChartMetaData &= "#DIFFICULTY 3" & vbCrLf
                                    Case &H12000
                                        ChartName = " [DP basic]"
                                        ChartMetaData &= "#DIFFICULTY 2" & vbCrLf
                                    Case &H16000
                                        ChartName = " [DP another]"
                                        ChartMetaData &= "#DIFFICULTY 4" & vbCrLf
                                    Case Else
                                        ChartName = "chart" & DataPadString(Hex(ChartOffset \ SkipAmount), 2, "0")
                                End Select
                                If bFoundTagInfo Then
                                    With ChartDBInfo
                                        ChartMetaData &= "#TITLE " & .Title & vbCrLf
                                        ChartMetaData &= "#ARTIST " & .Artist & vbCrLf
                                        ChartMetaData &= "#GENRE " & .Genre
                                    End With
                                End If
                                If Left(ChartName, 1) = " " Then
                                    If bFoundTagInfo Then
                                        ChartMetaData &= ChartName
                                        ChartName = DataGetFormattedName(ChartDBInfo.Title) & ChartName
                                    Else
                                        ChartName = "chart" & ChartName
                                    End If
                                End If

                                If ThisJob.ConvertChart Then
                                    sDecoderInfo = "Converting Chart +" & Hex(y)
                                    If bFoundTagInfo Then
                                        sDecoderInfo &= " [" & ChartDBInfo.Title & "]"
                                    End If
                                    ChartLoadMemory(ChartInfo, ChartTypes.IIDXCS5)
                                    If Chart.NoteCount > 0 Then
                                        Chart.AdjustAllNotes(1)
                                        If BGMNumber > 0 Then
                                            Chart.ConvertAllNotes(BGMNumber, 1)
                                        End If

                                        If KeysoundSets > 1 Then
                                            If Chart.NoteCount(2) > 0 Then
                                                ChartSaveBMS(ThisDataFolder & ChartName & ".bme", ChartTypes.IIDXCS5, 192, ChartMetaData, BGMFile, Prefix)
                                            Else
                                                ChartSaveBMS(ThisDataFolder & ChartName & ".bme", ChartTypes.IIDXCS5, 192, ChartMetaData, BGMFile)
                                            End If
                                        Else
                                            ChartSaveBMS(ThisDataFolder & ChartName & ".bme", ChartTypes.IIDXCS5, 192, ChartMetaData, BGMFile)
                                        End If
                                    End If
                                Else
                                    DataDumpArray(ThisDataFolder & ChartName & ".cs5", ChartInfo)
                                End If

                                Exit For
                            ElseIf (rbytes(x) = rbytes(x + 1) And rbytes(x + 1) = rbytes(x + 2) And rbytes(x + 2) = rbytes(x + 3)) And (rbytes(x) = &HA Or rbytes(x) = &H88 Or rbytes(x) = &H4F) Then
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next
        Next
        ChartInfo = Nothing
        AudioConvert = Nothing
        AudioData = Nothing
        ConvertedSample = Nothing
        sampleinfo = Nothing
        LastSampleInfo = Nothing
    End Sub

    Public Function DataDetectFormat(ByVal InputStream As IO.Stream, Optional ByVal FileOffset As Long = -1) As DataDetectedType
        Dim TableVars() As Integer
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim BGMcheck As Integer
        Dim dt As SoundDecodeType
        Dim dp As Integer
        Dim ThisPan As Integer
        Dim BaseOffset As Long
        ReDim TableVars(0 To 15)

        SoundRipInfo.DoRip = False
        DataDetectFormat = DataDetectedType.None

        'we need to be able to seek in format detection...
        If Not InputStream.CanSeek Then
            Exit Function
        End If

        'initialize stream binary reader
        If FileOffset >= 0 Then
            InputStream.Position = FileOffset
        Else
            FileOffset = InputStream.Position
        End If
        Dim Reader As New IO.BinaryReader(InputStream)
        For x = 0 To UBound(TableVars)
            TableVars(x) = Reader.ReadInt32
        Next
        InputStream.Position = FileOffset

        'movies
        If TableVars(0) = &HBA010000 Then
            DataDetectFormat = DataDetectedType.VOB
            Exit Function
        End If

        'CS charts
        If TableVars(0) = &H80000864 Then
            DataDetectFormat = DataDetectedType.IIDXNewChart
            Exit Function
        End If

        'newer DDR for Playstation 2 song, 9th+ IIDX BGMs
        If TableVars(0) = &H8640001 And TableVars(1) = 0 And TableVars(2) >= 48 Then
            DataDetectFormat = DataDetectedType.DDRPS2
            With SoundRipInfo
                .DoRip = True
                '.FileNumber = FileNumber
                .FileStream = InputStream
                .InstanceCount = 1
                ReDim .Instances(0)
            End With
            With SoundRipInfo.Instances(0)
                .FileOffset = FileOffset + TableVars(2)
                .MaxLength = TableVars(3)
                .Freq = TableVars(6)
                .LoopStart = TableVars(4)
                .LoopEnd = TableVars(5)
                .Channels = TableVars(7)
                .BlockSize = TableVars(9)
                .VolumeL = CSng(TableVars(10)) / 100
                .VolumeR = .VolumeL
                'volume boost
                .VolumeL *= 1.2
                .VolumeR *= 1.2
                .OutputName = ""
                dp = 0
                If TableVars(8) <> 0 Then
                    'FileGet(FileNumber, dp, FileOffset + &H801&)
                    InputStream.Position = FileOffset + &H800
                    dp = Reader.ReadInt32
                    .DecodeParam = dp
                End If
                .DecodeType = IIf(dp = 0, SoundDecodeType.IIDX9, SoundDecodeType.BMDXgeneric)
            End With
            Exit Function
        End If

        '9th+ Keysound sets
        Select Case (TableVars(0) And &HFF00FFFF)
            Case &H7077, &H6665, &H2CF
                DataDetectFormat = DataDetectedType.IIDXKeysound
                dt = SoundDecodeType.IIDX9
                With SoundRipInfo
                    .DoRip = True
                    .FileStream = InputStream
                    .InstanceCount = 0
                    ReDim .Instances(0)
                End With
                Dim HeadA11 As DataSampleHeaderA11
                Dim HeadB11 As DataSampleHeaderB11
                Dim InfoA11() As DataSampleInfoA11
                Dim InfoB11() As DataSampleInfoB11
                'read header
                InputStream.Position = FileOffset
                HeadA11.Ident = Reader.ReadInt16
                InputStream.Position = FileOffset + &H8000
                HeadB11.SampleCount = Reader.ReadInt32
                HeadB11.TotalLength = Reader.ReadInt32
                ReDim InfoA11(0 To 2046)
                ReDim InfoB11(0 To HeadB11.SampleCount - 1)
                If HeadB11.SampleCount > 0 Then

                    'read logical sample table
                    InputStream.Position = FileOffset + &H10
                    For x = 0 To UBound(InfoA11)
                        With InfoA11(x)
                            .Unk0 = Reader.ReadInt16
                            .unk1 = Reader.ReadByte
                            .ChanCount = Reader.ReadByte
                            .Unk2 = Reader.ReadInt32
                            .PanLeft = Reader.ReadByte
                            .PanRight = Reader.ReadByte
                            .SampleNum = Reader.ReadInt16
                            .volume = Reader.ReadByte
                            .unk3 = Reader.ReadByte
                            .Unk4 = Reader.ReadInt16
                        End With
                    Next

                    'read physical sample table
                    InputStream.Position = FileOffset + &H8010
                    For x = 0 To UBound(InfoB11)
                        With InfoB11(x)
                            .SampOffset = Reader.ReadInt32
                            .SampLength = Reader.ReadInt32
                            .ChanCount = Reader.ReadInt16
                            .Frequ = Reader.ReadInt16
                            .Unk0 = Reader.ReadInt32
                        End With
                    Next

                    'read decryption key (if any)
                    InputStream.Position = FileOffset + &H8010 + (HeadB11.SampleCount * 16)
                    dp = Reader.ReadInt32
                    dt = IIf(dp = 0, SoundDecodeType.IIDX9, SoundDecodeType.BMDXgeneric)
                    BaseOffset = &H8020& + (HeadB11.SampleCount * 16) + FileOffset

                    'translate sample table to rip parameters
                    x = 1
                    If InfoA11(1).SampleNum <> InfoA11(0).SampleNum Then
                        x = 0
                    End If
                    ReDim SoundRipInfo.Instances(0 To 2046)
                    y = 0
                    Do Until x = 2046
                        z = InfoA11(x).SampleNum
                        If InfoA11(x).ChanCount > 0 And InfoB11(z).ChanCount > 0 Then
                            With SoundRipInfo.Instances(y)
                                .BlockSize = InfoB11(z).SampLength
                                .Channels = InfoB11(z).ChanCount
                                .DecodeParam = dp
                                .DecodeType = dt
                                .FileOffset = InfoB11(z).SampOffset + BaseOffset
                                .Freq = SoundConvertFrequency(InfoB11(z).Frequ)
                                .LoopEnd = 0
                                .LoopStart = 0
                                .MaxLength = .BlockSize * .Channels
                                .OutputName = DataKeySoundFileName(x + 1) 'DataPadString(CStr(x + 1), 4, "0") 'change this later to BME numbering menu selection
                                .Pan = 0.5
                                If .Channels = 1 Then
                                    ThisPan = (InfoA11(x).PanLeft + InfoA11(x).PanRight) / 2
                                    .VolumeL = PanTableRight(Int((ThisPan / 128) * 256))
                                    .VolumeR = PanTableLeft(Int((ThisPan / 128) * 256))
                                ElseIf .Channels = 2 Then
                                    If InfoA11(x).PanLeft <= InfoA11(x).PanRight Then
                                        .VolumeL = PanTableRight((InfoA11(x).PanLeft / 128) * 256)
                                        .VolumeR = PanTableLeft((InfoA11(x).PanRight / 128) * 256)
                                    Else
                                        .VolumeL = PanTableLeft((InfoA11(x).PanLeft / 128) * 256)
                                        .VolumeR = PanTableRight((InfoA11(x).PanRight / 128) * 256)
                                    End If
                                Else
                                    .VolumeL = 0
                                    .VolumeR = 0
                                End If
                                .VolumeL *= (InfoA11(x).volume / 128)
                                .VolumeR *= (InfoA11(x).volume / 128)
                            End With
                            y += 1
                        End If
                        x += 1
                    Loop
                End If
                SoundRipInfo.InstanceCount = y
                ReDim Preserve SoundRipInfo.Instances(0 To y - 1)
                Exit Function
        End Select

        'older PS2 BGM (probably strictly DDR)
        If TableVars(0) = &H67617653 Then
            DataDetectFormat = DataDetectedType.DDRPS2
            With SoundRipInfo
                .DoRip = True
                '.FileNumber = FileNumber
                .FileStream = InputStream
                .InstanceCount = 1
                ReDim .Instances(0)
            End With
            With SoundRipInfo.Instances(0)
                .DecodeType = SoundDecodeType.DDRPSX
                .FileOffset = FileOffset + &H800
                .Freq = TableVars(2)
                .LoopStart = 0
                .LoopEnd = 0
                .Channels = (TableVars(3) And &HFF)
                .BlockSize = &H2000
                .VolumeL = 1
                .VolumeR = 1
                .OutputName = ""
            End With
            Exit Function
        End If

        'older PS2 beatmaniaIIDX bgm
        If ((TableVars(0) And &HFF000000) = 0) And TableVars(2) = 2 And TableVars(3) = 0 And TableVars(4) = 0 And TableVars(5) = 0 Then
            InputStream.Position = FileOffset + &H800
            BGMcheck = Reader.ReadInt32
            If BGMcheck = &H200 Then
                DataDetectFormat = DataDetectedType.IIDXOldBGM
                With SoundRipInfo
                    .DoRip = True
                    .FileStream = InputStream
                    .InstanceCount = 1
                    ReDim .Instances(0)
                End With
                With SoundRipInfo.Instances(0)
                    .DecodeType = SoundDecodeType.DDRPS2
                    .FileOffset = FileOffset + &H800
                    .Freq = ((TableVars(1) >> 24) And &HFF) Or ((TableVars(1) >> 8) And &HFF00)
                    .LoopStart = 0
                    .LoopEnd = 0
                    .Channels = 2
                    .BlockSize = &H800
                    .VolumeL = ((TableVars(1) And &HFF00) >> 8) / 100
                    .VolumeR = ((TableVars(1) And &HFF00) >> 8) / 100
                    .OutputName = ""
                    .MaxLength = (TableVars(0) And &HFF)
                    .MaxLength <<= 8
                    .MaxLength += (TableVars(0) And &HFF00) >> 8
                    .MaxLength <<= 8
                    .MaxLength += (TableVars(0) And &HFF0000) >> 16
                    .MaxLength <<= 8
                    .MaxLength += (TableVars(0) And &HFF000000) >> 24
                End With
                Exit Function
            End If
        End If


        'Public Structure DataSampleInfo3
        '    Public SampleNum As Short
        '    Public Unk0 As Short
        '    Public unk1 As Byte
        '    Public vol As Byte
        '    Public pan As Byte
        '    Public SampType As Byte
        '    Public FreqLeft As Integer
        '    Public FreqRight As Integer
        '    Public OffsLeft As Integer
        '    Public OffsRight As Integer
        '    Public PseudoLeft As Integer
        '    Public PseudoRight As Integer
        'End Structure


        'older PS2 keysound tables
        If TableVars(0) > 0 And TableVars(1) > 0 And TableVars(2) >= 0 And TableVars(2) < TableVars(1) And TableVars(1) < TableVars(0) Then
            If ((TableVars(0) - TableVars(2)) - TableVars(1) = 16384) And TableVars(3) = 0 Then
                Dim Key3(510) As DataSampleInfo3
                ReDim SoundRipInfo.Instances(0 To 510)
                'FileSystem.FileGet(FileNumber, Key3, FileOffset + &H21)
                InputStream.Position = FileOffset + &H20
                For x = 0 To 510
                    With Key3(x)
                        .SampleNum = Reader.ReadInt16
                        .Unk0 = Reader.ReadInt16
                        .unk1 = Reader.ReadByte
                        .vol = Reader.ReadByte
                        .pan = Reader.ReadByte
                        .SampType = Reader.ReadByte
                        .FreqLeft = Reader.ReadInt32
                        .FreqRight = Reader.ReadInt32
                        .OffsLeft = Reader.ReadInt32
                        .OffsRight = Reader.ReadInt32
                        .PseudoLeft = Reader.ReadInt32
                        .PseudoRight = Reader.ReadInt32
                    End With
                Next

                y = -1
                For x = 0 To 510
                    With Key3(x)
                        If (.SampleNum > 0) And (.SampType > 0) And (.FreqLeft > 0) Then
                            y += 1
                            Select Case .SampType
                                Case 2, 4
                                    With SoundRipInfo.Instances(y)
                                        If Key3(x).SampType = 2 Then
                                            .BlockSize = 16
                                            .Channels = 1
                                        ElseIf Key3(x).SampType = 4 Then
                                            .BlockSize = Key3(x).OffsRight - Key3(x).OffsLeft
                                            .Channels = 2
                                        End If
                                        .FileOffset = (FileOffset + Key3(x).OffsLeft) - 61456
                                    End With
                                Case 3
                                    With SoundRipInfo.Instances(y)
                                        .BlockSize = 16
                                        .Channels = 1
                                        .FileOffset = TableVars(1) + FileOffset + Key3(x).OffsLeft + 16400
                                    End With
                            End Select
                            With SoundRipInfo.Instances(y)
                                .DecodeType = SoundDecodeType.IIDX3
                                .Freq = Key3(x).FreqLeft
                                .LoopStart = 0
                                .LoopEnd = 0
                                .MaxLength = -1
                                .OutputName = DataKeySoundFileName(Key3(x).SampleNum)
                                .VolumeL = Key3(x).vol / &H80
                                .VolumeR = Key3(x).vol / &H80
                                .VolumeL *= PanTableLeft(Key3(x).pan * 2)
                                .VolumeR *= PanTableRight(Key3(x).pan * 2)
                            End With
                        End If
                    End With
                Next
                If y >= 0 Then
                    ReDim Preserve SoundRipInfo.Instances(y)
                End If
                With SoundRipInfo
                    .DoRip = True
                    .FileStream = InputStream
                    .InstanceCount = y + 1
                End With
                DataDetectFormat = DataDetectedType.IIDXOldKeysound
            End If
        End If

        'PSX BGM (probably strictly DDR, also not for CAT files)
        If TableVars(0) = 0 And TableVars(1) = 0 And TableVars(2) = 0 And TableVars(3) = 0 And ((TableVars(4) And &H400) > 0) Then
            DataDetectFormat = DataDetectedType.DDRPSX
            With SoundRipInfo
                .DoRip = True
                .FileStream = InputStream
                .InstanceCount = 1
                ReDim .Instances(0)
            End With
            With SoundRipInfo.Instances(0)
                .FileOffset = FileOffset
                .Freq = 44100
                .LoopStart = 0
                .LoopEnd = 0
                .Channels = 2
                .BlockSize = &H4000
                .VolumeL = 1
                .VolumeR = 1
                .OutputName = ""
                .MaxLength = -1
            End With
            Exit Function
        End If

        'PSX CAT BGM (commonly found in Disney and other special Dancing Stage editions)
        If TableVars(0) = 0 And TableVars(1) = 0 And TableVars(2) = 0 And TableVars(3) = 0 And ((TableVars(4) And &HFF) <> 0) And ((TableVars(4) And &HF000) = 0) Then
            DataDetectFormat = DataDetectedType.DDRPSX
            With SoundRipInfo
                .DoRip = True
                .FileStream = InputStream
                .InstanceCount = 1
                ReDim .Instances(0)
            End With
            With SoundRipInfo.Instances(0)
                .FileOffset = FileOffset
                .Freq = 44100
                .LoopStart = 0
                .LoopEnd = 0
                .Channels = 2
                .BlockSize = &H4000
                .VolumeL = 1
                .VolumeR = 1
                .OutputName = ""
                .MaxLength = -1
            End With
            Exit Function
        End If

        'CS2 charts
        If (TableVars(0) And &HFFFF03) = &H800002 Then
            DataDetectFormat = DataDetectedType.IIDXOldChart
            Exit Function
        End If

        'Pop'n Music 11+ charts (CS)
        If (TableVars(0) = &H20) And (TableVars(8) = 0) And (TableVars(9) <> 0) Then

        End If

        'VAG standard Playstation audio header (values are MSB first)
        '0    4    VAG? (? being 'i' or 'p')
        '4    4    blocksize
        '8    4    
        'C    4    length in bytes
        '10   4    frequency in Hz
        '14   4    loop start
        '18   4    loop end
        '1C   4
        '20   16   name
        If (TableVars(0) And &HFFFFFF) = &H474156 Then
            Select Case (TableVars(0) >> 24)
                Case &H70 'p'
                    DataDetectFormat = DataDetectedType.VAG
                    With SoundRipInfo
                        .DoRip = True
                        .FileStream = InputStream
                        .InstanceCount = 1
                        ReDim .Instances(0)
                    End With
                    With SoundRipInfo.Instances(0)
                        .FileOffset = FileOffset + &H30
                        .Freq = DataSwapInt32(TableVars(4))
                        .LoopStart = DataSwapInt32(TableVars(5))
                        .LoopEnd = DataSwapInt32(TableVars(6))
                        .BlockSize = DataSwapInt32(TableVars(1))
                        .Channels = IIf(.BlockSize > 0, 2, 1)
                        .VolumeL = 1
                        .VolumeR = 1
                        .OutputName = ""
                        .MaxLength = DataSwapInt32(TableVars(3))
                    End With
                    Exit Function
                Case &H69 'interleave
            End Select
        End If

        If DataDetectFormat = DataDetectedType.None Then
            DebugLog("DataDetectFormat", "Unknown format.", "The format for this file can't be determined." & vbCrLf & "File offset: " & FileOffset)
        End If


    End Function

    'encode the Bemani LZ format
    Public Function DataEncodeBemaniLZ77(ByRef InData() As Byte, ByRef OutData() As Byte) As Integer
        Dim EncStream As New IO.MemoryStream
        Dim DecBytes() As Byte
        Dim LineData(0 To 16) As Byte
        ReDim DecBytes(0 To UBound(InData))
        Array.Copy(InData, 0, DecBytes, 0, UBound(InData) + 1)
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim e As Boolean = False
        Dim ThisCode As Byte = 0
        Dim LineLength As Byte = 1
        Dim CompressType As Byte = 0
        Dim CompressLength As Integer = 0
        Dim CompressOffset As Integer = 0
        Dim bCompress As Boolean = False
        Dim SearchLength As Integer = 0
        Dim SearchOffset As Integer = 0
        Dim SearchOffsetDiff As Integer = 0
        x = 0
        Do
            SearchLength = Math.Min(34, (UBound(DecBytes) - x) + 1)
            bCompress = False
            CompressType = 0
            Do While SearchLength > 2
                For SearchOffset = x - SearchLength To 0 Step -1
                    If SearchOffset < 0 Then
                        Exit For
                    End If
                    SearchOffsetDiff = x - SearchOffset
                    If SearchOffsetDiff > 1023 Then
                        Exit For
                    End If

                    'short jump?
                    If (SearchLength >= 2 And SearchLength <= 5) And (SearchOffsetDiff >= 1 And SearchOffsetDiff <= 16) Then
                        e = True
                        For y = 0 To SearchLength - 1
                            If DecBytes(x + y) <> DecBytes(SearchOffset + y) Then
                                e = False
                                Exit For
                            End If
                        Next
                        If e Then
                            bCompress = True
                            CompressType = 2
                            CompressOffset = SearchOffsetDiff
                            CompressLength = SearchLength
                            bCompress = True
                            Exit Do
                        End If
                    End If

                    'long jump?
                    If (SearchLength >= 3 And SearchLength <= 34) And (SearchOffsetDiff >= 0 And SearchOffsetDiff <= 1023) Then
                        e = True
                        For y = 0 To SearchLength - 1
                            If DecBytes(x + y) <> DecBytes(SearchOffset + y) Then
                                e = False
                                Exit For
                            End If
                        Next
                        If e Then
                            bCompress = True
                            CompressType = 1
                            CompressOffset = SearchOffsetDiff
                            CompressLength = SearchLength
                            bCompress = True
                            Exit Do
                        End If
                    End If

                Next
                SearchLength -= 1
            Loop
            Select Case CompressType
                Case 0 'none (direct copy)
                    LineData(LineLength) = DecBytes(x)
                    LineLength += 1
                    x += 1
                    bCompress = False
                Case 1 'long jump
                    LineData(LineLength) = (CompressLength - 3) << 2
                    LineData(LineLength) += (CompressOffset >> 8)
                    LineData(LineLength + 1) = (CompressOffset And 255)
                    LineLength += 2
                    x += CompressLength
                Case 2 'short jump
                    LineData(LineLength) = &H80
                    LineData(LineLength) += ((CompressLength - 2) << 4)
                    LineData(LineLength) += ((CompressOffset - 1) And 15)
                    LineLength += 1
                    x += CompressLength
                    'Case 3 'uncompressable blocks
            End Select
            If bCompress Then
                LineData(0) = LineData(0) Or (2 ^ ThisCode)
            End If
            ThisCode += 1
            If ThisCode = 8 Then
                ThisCode = 0
                EncStream.Write(LineData, 0, LineLength)
                LineLength = 1
                LineData(0) = 0
            End If
            If x > UBound(DecBytes) Then
                LineData(0) = LineData(0) Or (2 ^ ThisCode)
                LineData(LineLength) = &HFF
                LineLength += 1
                EncStream.Write(LineData, 0, LineLength)
                Exit Do
            End If
        Loop
        OutData = EncStream.ToArray()
        EncStream.Close()
        Return UBound(OutData) + 1
    End Function

    'todo: use streams instead of expanding the output array
    'decode the Bemani LZ format
    Public Function DataDecodeBemaniLZ77(ByRef inData() As Byte, ByRef outData() As Byte) As Integer
        Dim BytesDecoded As Integer = 0
        Dim EndFlag As Boolean = False
        Dim DecodeOffs As Integer = 0
        Dim DecLength As Integer = 0
        Dim dec() As Byte
        Dim i As Integer = 0
        Dim flags As Integer = 0
        Dim j As Integer = 0
        Dim c As Byte = 0
        Dim xloop As Boolean = False
        ReDim dec(0)
        Do
            Do
                flags >>= 1
                If (flags And &H100) = 0 Then
                    flags = (inData(DecodeOffs) Or &HFF00)
                    DecodeOffs += 1
                End If
                c = inData(DecodeOffs)
                If (flags And 1) = 0 Then
                    DataAppendByteToArray(dec, c)
                    DecodeOffs += 1
                    Exit Do
                End If
                'i is offset
                'j is count
                If (c And &H80) = 0 Then
                    i = inData(DecodeOffs + 1)
                    DecodeOffs += 2
                    i = i Or ((c And 3) << 8)
                    j = (c >> 2) + 2
                    xloop = True
                End If
                If (Not xloop) Then
                    DecodeOffs += 1
                    If (c And &H40) = 0 Then
                        i = (c And 15) + 1
                        j = (c >> 4) + 1 - 8 '-8 gets rid of the identifier bit
                        xloop = True
                    End If
                End If
                If xloop Then
                    xloop = False
                    Do
                        DecLength = UBound(dec)
                        If ((DecLength + 1) - i > 0) And ((DecLength + 1) - i < DecLength) Then
                            DataAppendByteToArray(dec, dec((DecLength + 1) - i))
                        ElseIf ((DecLength + 1) - i = DecLength) Then
                            DataAppendByteToArray(dec, dec(DecLength))
                        Else
                            DataAppendByteToArray(dec, 0)
                        End If
                        j -= 1
                    Loop While j >= 0
                    Exit Do
                End If
                If c = 255 Then
                    EndFlag = True
                    Exit Do
                End If
                j = c - &HC0 + 7
                Do
                    DataAppendByteToArray(dec, inData(DecodeOffs))
                    DecodeOffs += 1
                    j -= 1
                Loop While j >= 0
            Loop
        Loop While EndFlag = False
        BytesDecoded = UBound(dec)
        ReDim outData(BytesDecoded - 1)
        Array.Copy(dec, 1, outData, 0, BytesDecoded)
        Return BytesDecoded
    End Function

    Public Sub DataAppendByteToArray(ByRef arr() As Byte, ByVal val As Byte)
        Dim a As Integer = UBound(arr)
        ReDim Preserve arr(a + 1)
        arr(a + 1) = val
    End Sub

    Public Function DataBMEString(ByVal i As Integer, Optional ByVal pad As Integer = 2, Optional ByVal BMEString As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ") As String
        Do While i < 0
            i += Len(BMEString) ^ pad
        Loop
        Dim ThisValue As Integer = Len(BMEString)
        DataBMEString = String.Empty
        Do While Len(DataBMEString) < pad
            DataBMEString = Mid(BMEString, (i Mod ThisValue) + 1, 1) & DataBMEString
            i \= ThisValue
        Loop
    End Function

    Public Function DataUnBMEString(ByVal s As String, Optional ByVal BMEString As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ") As Integer
        Dim ThisValue As Integer = Len(BMEString)
        Dim i As Integer
        DataUnBMEString = 0
        For i = 1 To Len(s)
            DataUnBMEString *= ThisValue
            DataUnBMEString += InStr(BMEString, Mid(s, i, 1)) - 1
        Next
    End Function

    Public Function DataKeySoundFileName(ByVal i As Integer) As String
        Select Case DataNamingStyle
            Case 0 : Return DataBMEString(i, 2) 'bme
            Case 1 : Return DataBMEString(i, 2, "0123456789ABCDEF") 'bms
            Case 2 : Return DataBMEString(i, 4) 'xbme
            Case 3 : Return CStr(i) 'decimal
            Case 4 : Return DataBMEString(i, 2, "0123456789") 'fixed decimal
            Case 5 : Return DataBMEString(i, 4, "0123456789") 'extended decimal
            Case Else : Return CStr(i) 'return decimal by default
        End Select
    End Function

    Public Function DataSwapInt32(ByVal i As Int32) As Int32
        Dim x As Integer
        DataSwapInt32 = 0
        For x = 0 To 3
            DataSwapInt32 <<= 8
            DataSwapInt32 = DataSwapInt32 Or (i And 255)
            i >>= 8
        Next
    End Function

    Public Function DataMakeInt32(Optional ByVal b1 As Byte = 0, Optional ByVal b2 As Byte = 0, Optional ByVal b3 As Byte = 0, Optional ByVal b4 As Byte = 0) As Int32
        DataMakeInt32 = b4
        DataMakeInt32 <<= 8
        DataMakeInt32 += b3
        DataMakeInt32 <<= 8
        DataMakeInt32 += b2
        DataMakeInt32 <<= 8
        DataMakeInt32 += b1
    End Function

    Public Function DataMakeInt64(Optional ByVal b1 As Byte = 0, Optional ByVal b2 As Byte = 0, Optional ByVal b3 As Byte = 0, Optional ByVal b4 As Byte = 0, Optional ByVal b5 As Byte = 0, Optional ByVal b6 As Byte = 0, Optional ByVal b7 As Byte = 0, Optional ByVal b8 As Byte = 0) As Int64
        DataMakeInt64 = b8
        DataMakeInt64 <<= 8
        DataMakeInt64 += b7
        DataMakeInt64 <<= 8
        DataMakeInt64 += b6
        DataMakeInt64 <<= 8
        DataMakeInt64 += b5
        DataMakeInt64 <<= 8
        DataMakeInt64 += b4
        DataMakeInt64 <<= 8
        DataMakeInt64 += b3
        DataMakeInt64 <<= 8
        DataMakeInt64 += b2
        DataMakeInt64 <<= 8
        DataMakeInt64 += b1
    End Function

End Module
