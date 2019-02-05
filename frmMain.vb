Option Explicit On
Imports System.Threading

Public Class frmMain

    Private MainTimer As New clsHiResTimer
    Private sChartPlayback As String
    Private bProgramLoaded As Boolean
    Private sDrives() As String
    Private bUseSimpleMode As Boolean
    Private bWorking As Boolean
    Private bDecoderState As Boolean
    Private bExtraFile As Boolean
    Private bConvertSound As Boolean
    Private bConvertChart As Boolean
    Private bUseCHD As Boolean
    Private LogCount As Integer = -1
    Private bStopProcess As Boolean = False
    Private bPlaying As Boolean
    Private bStopPlaying As Boolean
    Private bShutdown As Boolean
    Private ThisTime As Integer
    Private ThisEvent As Integer
    Private ThisMetric As Double
    Private ChartFile As String
    Private PlaybackThread As Thread
    Private WindowHandle As IntPtr
    Private bQueueActive As Boolean
    Private AutoNameTable As DataTable

    Private TimingDefaults() As String = {"Troopers/Popn (1ms)", "1", "IIDX14 GOLD (60.04tps)", "16.656", "IIDX9 AC (59.94tps)", "16.683", "IIDX firebeat (59.8tps)", "16.722", "IIDX alt firebeat (59.82tps)", "16.717", "beatmania (58tps)", "17.241", "Pop'n (50/3tps)", "16.667"}
    Private ChartFormats() As String = {"IIDXAC (.1)", ChartTypes.IIDXAC, "IIDX 4-byte (.cs2)", ChartTypes.IIDXCS2, "IIDX 8-byte (.cs)", ChartTypes.IIDXCS, "beatmania (.cs5)", ChartTypes.IIDXCS5, "BMS/BME", ChartTypes.BME, "PMS", ChartTypes.PMS}




    Private Sub QueueRip(ByVal bPriority As Boolean)
        Dim NewJob As QueueJob
        Dim ConfigIndexes() As Integer
        Dim ConfigIndexCount As Integer = 1
        Dim x As Integer
        ReDim ConfigIndexes(0)
        bOneClickMode = True
        If sTargetFolder = "" Then
            If bUseSimpleMode Then
                TabControl1.SelectTab(tabSimple)
            Else
                TabControl1.SelectTab(tabAdvanced)
            End If
            MsgBox("Please specify a Target Folder where the extracted content will be saved to.", MsgBoxStyle.Information)
            Exit Sub
        End If
        If bUseSimpleMode Then
            ConfigIndexCount = 0
            For x = 0 To UBound(Config)
                With Config(x)
                    If .GameID = cmbSimpleGame.SelectedItem Then
                        ReDim Preserve ConfigIndexes(ConfigIndexCount)
                        ConfigIndexes(ConfigIndexCount) = x
                        ConfigIndexCount += 1
                    End If
                End With
            Next
        Else
            ConfigIndexes(0) = cmbAdvancedGame.SelectedIndex
        End If

        For x = 0 To UBound(ConfigIndexes)
            cmbAdvancedGame.SelectedIndex = ConfigIndexes(x)
            bExtraFile = (chkAdvancedUseExtra.Checked = True) And (sExtraFile <> "")
            If sSourceFile = "" Then
                If bUseSimpleMode Then
                    MsgBox("Please specify the game to extract media from.", MsgBoxStyle.Information)
                    TabControl1.SelectTab(tabSimple)
                Else
                    MsgBox("Please specify a Source File where the content will be extracted from.", MsgBoxStyle.Information)
                    TabControl1.SelectTab(tabAdvanced)
                End If
                Exit Sub
            End If
            If sExtraFile = "" And chkAdvancedUseExtra.Checked = True Then
                TabControl1.SelectTab(tabAdvanced)
                MsgBox("Please specify the Extra Info file to be used. Or, uncheck the box to not use one.", MsgBoxStyle.Information)
                Exit Sub
            End If
            If Strings.Right(sTargetFolder, 1) <> "\" Then
                sTargetFolder += "\"
            End If
            With NewJob
                .SourceFile = sSourceFile
                If bExtraFile Then
                    .SourceExtraFile = sExtraFile
                Else
                    .SourceExtraFile = ""
                End If
                .TargetFolder = sTargetFolder
                .AutoName = chkAutoName.Checked
                .AutoStructure = chkAutoPair.Checked
                .ConvertBGM = chkConvertSounds.Checked
                .ConvertKeysounds = chkConvertSounds.Checked
                .ConvertChart = chkConvertCharts.Checked
                .DontDecompress = chkExtractOnly.Checked
                .GameConfig = iGameType
                .Name = Config(iGameType).Name & " (" & Config(iGameType).File & ")"
                .RipBGM = chkRipBGM.Checked
                .RipCharts = chkRipCharts.Checked
                .RipGraphics = chkRipGraphics.Checked
                .RipKeysounds = chkRipKeysounds.Checked
                .RipVideos = chkRipVideos.Checked
                .StripSilence = chkStripSilence.Checked
            End With
            If bPriority Then
                QueueAddJobPriority(NewJob)
                lstQueue.Items.Insert(0, NewJob.Name)
            Else
                QueueAddJob(NewJob)
                lstQueue.Items.Add(NewJob.Name)
            End If
        Next
    End Sub


    Private Sub Rip()
        If (Not bWorking) Then
            bWorking = True
            FileCreateFolder(sTargetFolder)
            DataNamingStyle = cmbKeysoundNumbering.SelectedIndex
            bConvertChart = chkConvertCharts.Checked
            bConvertSound = chkConvertSounds.Checked
            ConfigLoadFormats()
            MainDecoder.RunWorkerAsync()
        End If
    End Sub








    Private Sub MainDecoder_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles MainDecoder.DoWork
        Dim SourceStream As IO.Stream = New IO.MemoryStream
        Dim ExtraStream As IO.Stream = New IO.MemoryStream
        Dim ThisStream As IO.Stream
        If (Not My.Computer.FileSystem.FileExists(ThisJob.SourceFile)) Then
            bWorking = False
            Exit Sub
        Else
            SourceStream = New IO.FileStream(ThisJob.SourceFile, IO.FileMode.Open, IO.FileAccess.Read)
            If (ThisJob.SourceExtraFile <> "") AndAlso (My.Computer.FileSystem.FileExists(ThisJob.SourceExtraFile)) Then
                ExtraStream = New IO.FileStream(ThisJob.SourceExtraFile, IO.FileMode.Open, IO.FileAccess.Read)
            End If
        End If

        Dim ThisFileNumber As Integer
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim o As Long = Val(txtAdvancedOffset.Text)
        Dim oi As Long = Val(txtAdvancedInterval.Text)
        Dim BGMFolder As String
        Dim KeyFolder As String
        Dim ChartFolder As String
        Dim DataFileEntry As DataFileEntry
        Dim FormattedName As String
        Dim ol As Long = 0
        Dim dec() As Byte = {}
        Dim si As Integer = 0
        Dim sf As String
        Dim CHDStartOffset As Integer = 0
        Dim sChartString As String
        Dim sWave01Override As String
        Dim DecFormat As DataDetectedType
        bWorking = True
        If SourceStream.Length < 1 Then
            'could not open the source file...
            bWorking = False
            Exit Sub
        End If

        If (ThisJob.SourceExtraFile <> "") And (Config(ThisJob.GameConfig).CHD = False) Then
            sDecoderInfo = "Reading File Table"
            If DataReadFileTable(ThisJob.SourceExtraFile, ThisJob.GameConfig) Then
                'there was a problem reading the file table, exit...
                bWorking = False
                Exit Sub
            End If
            FileCreateFolder(ThisJob.TargetFolder)
            x = 0
            'fi = FileLoad(ThisJob.SourceExtraFile)
            For x = 0 To UBound(DataFiles)
                DataFileEntry = DataFiles(x)
                sDecoderInfo = "Parsing File Table (" & DataFileEntry.Offset & ")"
                SongString = DataPadString(CStr(x + Config(ThisJob.GameConfig).Index), 4, "0") & " "
                If DataFileEntry.IsFromInfoFile Then
                    'ThisFileNumber = fi
                    ThisStream = ExtraStream
                Else
                    'ThisFileNumber = f
                    ThisStream = SourceStream
                End If
                If ThisJob.DontDecompress Then
                    With DataFileEntry
                        If .Length > 0 Then
                            'DataExtractRaw(ThisFileNumber, ThisJob.TargetFolder & Trim(SongString & DataGetFormattedName(.Name)) & ".raw", .Offset, .Length)
                            DataExtractRaw(ThisStream, ThisJob.TargetFolder & Trim(SongString & DataGetFormattedName(.Name)) & ".raw", .Offset, .Length)
                        End If
                    End With
                Else
                    With DataFileEntry
                        If (.IsFromInfoFile = True) And (.ForcedType <> DataDetectedType.None) Then
                            DecFormat = .ForcedType
                        Else
                            DecFormat = DataDetectFormat(ThisStream, .Offset)
                        End If
                        FormattedName = DataGetFormattedName(.Name)
                        If bOneClickMode Then
                            BGMFolder = ThisJob.TargetFolder & FormattedName & "\"
                            KeyFolder = ThisJob.TargetFolder & FormattedName & "\"
                            ChartFolder = ThisJob.TargetFolder & FormattedName & "\"
                        Else
                            BGMFolder = ThisJob.TargetFolder
                            KeyFolder = ThisJob.TargetFolder & SongString & FormattedName & "\"
                            ChartFolder = ThisJob.TargetFolder
                        End If
                        Select Case DecFormat
                            Case DataDetectedType.DDRPS2, DataDetectedType.IIDXOldBGM
                                If chkRipBGM.Checked Then
                                    FileCreateFolder(BGMFolder)
                                    If chkConvertSounds.Checked Then
                                        sDecoderInfo = "Decoding BGM: " & .Name
                                        SoundStartDecode(ThisJob.SourceFile, BGMFolder & SongString & FormattedName, .Prefix)
                                    Else
                                        sDecoderInfo = "Extracting BGM: " & .Name
                                        DataExtractRaw(ThisFileNumber, BGMFolder & SongString & DataGetFormattedName(.Name) & ".bgm.raw", .Offset, .Length)
                                    End If
                                End If
                            Case DataDetectedType.IIDXKeysound, DataDetectedType.IIDXOldKeysound
                                If chkRipKeysounds.Checked Then
                                    If chkConvertSounds.Checked Then
                                        sDecoderInfo = "Decoding Keysounds: " & .Name
                                        FileCreateFolder(KeyFolder)
                                        SoundStartDecode(ThisJob.SourceFile, KeyFolder, .Prefix)
                                    Else
                                        sDecoderInfo = "Extracting Keysounds: " & .Name
                                        DataExtractRaw(ThisFileNumber, BGMFolder & SongString & DataGetFormattedName(.Name) & ".keys.raw", .Offset, .Length)
                                    End If
                                End If
                            Case DataDetectedType.IIDXNewChart, DataDetectedType.IIDXOldChart
                                If chkRipCharts.Checked Then
                                    ChartForceTiming = .ForceTiming
                                    sDecoderInfo = "Decoding Chart: " & .Name
                                    If DecFormat = DataDetectedType.IIDXOldChart Then
                                        ChartLoadFromOpen(ThisFileNumber, .Offset, DataEncodingType.KonamiLZ77, ChartTypes.IIDXCS2, .Length, (chkConvertCharts.Checked = False))
                                    Else
                                        ChartLoadFromOpen(ThisFileNumber, .Offset, DataEncodingType.KonamiLZ77, ChartTypes.IIDXCS, .Length, (chkConvertCharts.Checked = False))
                                    End If
                                    FileCreateFolder(ChartFolder)
                                    If chkConvertCharts.Checked Then
                                        sChartString = ""
                                        sWave01Override = ""
                                        If .Movie > 0 Then
                                            sChartString &= "#VIDEOFILE " & DataPadString(CStr(.Movie), 4, "0") & ".vob" & vbCrLf
                                        End If
                                        If .Keysound > 0 And .BGM > 0 Then
                                            sWave01Override = DataPadString(CStr(.BGM), 4, "0") & " " & DataGetFormattedName(DataFiles(.BGM).Name) & ".wav"
                                        End If
                                        If bChartInfo Then
                                            '.Genre = "genre"
                                            If .Genre = "" Then
                                                sChartString &= "#TITLE " & .Name & .Suffix & vbCrLf
                                            Else
                                                sChartString &= "#TITLE " & .Name & vbCrLf
                                            End If
                                            If .Artist <> "" Then
                                                sChartString &= "#ARTIST " & .Artist & vbCrLf
                                            End If
                                            If .Genre <> "" Then
                                                sChartString &= "#GENRE " & .Genre & .Suffix & vbCrLf
                                            End If
                                            If .Difficulty > 0 Then
                                                sChartString &= "#PLAYLEVEL " & .Difficulty & vbCrLf
                                            End If
                                            Select Case Trim(UCase(.Suffix))
                                                Case "[BEGINNER7]", "[BEGINNER14]"
                                                    sChartString &= "#DIFFICULTY 1" & vbCrLf
                                                Case "[LIGHT7]", "[LIGHT14]", "[NORMAL7]", "[NORMAL14]"
                                                    sChartString &= "#DIFFICULTY 2" & vbCrLf
                                                Case "[7KEY]", "[14KEY]", "[HYPER7]", "[HYPER14]", "[5KEY]", "[10KEY]"
                                                    sChartString &= "#DIFFICULTY 3" & vbCrLf
                                                Case "[ANOTHER7]", "[ANOTHER14]"
                                                    sChartString &= "#DIFFICULTY 4" & vbCrLf
                                                Case Else
                                                    x = x
                                            End Select
                                            ChartSaveBMS(ChartFolder & DataGetFormattedName(.Name) & .Suffix & ".bme", DataNamingStyle, 192, sChartString, sWave01Override, DataFiles(.Keysound).Prefix)
                                        Else
                                            ChartSaveBMS(ChartFolder & SongString & DataGetFormattedName(.Name) & ".bme", DataNamingStyle, 192)
                                        End If
                                    Else
                                        Select Case DecFormat
                                            Case DataDetectedType.IIDXOldChart
                                                My.Computer.FileSystem.WriteAllBytes(ChartFolder & SongString & DataGetFormattedName(.Name) & .Suffix & ".cs2", ChartRaw, False)
                                            Case Else
                                                My.Computer.FileSystem.WriteAllBytes(ChartFolder & SongString & DataGetFormattedName(.Name) & .Suffix & ".cs", ChartRaw, False)
                                        End Select

                                    End If
                                End If
                            Case DataDetectedType.VOB
                                If chkRipVideos.Checked Then
                                    sDecoderInfo = "Extracting Video: " & SongString & .Name
                                    DataExtractRaw(ThisFileNumber, ThisJob.TargetFolder & Trim(SongString) & ".vob", .Offset, .Length)
                                End If
                        End Select
                    End With
                End If
                DataFileProgress = 100 * (x / UBound(DataFiles))
                If (bStopProcess) Or (bShutdown) Then
                    Exit For
                End If
            Next
        ElseIf (ThisJob.SourceExtraFile <> "") And (Config(ThisJob.GameConfig).CHD = False) Then
            FileCreateFolder(ThisJob.TargetFolder)
            ol = SourceStream.Length 'LOF(f)
            sDecoderInfo = "Scanning for Media"
            Do While (o < ol) And (bStopProcess = False) And (bShutdown = False)
                sf = ThisJob.TargetFolder & DataPadString(CStr(si), 4, "0") & " "
                Select Case DataDetectFormat(SourceStream, o)
                    Case DataDetectedType.DDRPS2
                        sDecoderInfo = "Writing: " & sf
                        SoundStartDecode(ThisJob.SourceFile, sf)
                        si += 1
                    Case DataDetectedType.IIDXKeysound
                        sDecoderInfo = "Writing: " & sf
                        SoundStartDecode(ThisJob.SourceFile, sf)
                        si += 1
                End Select
                sDecoderInfo = "Scanning for Media"
                o += oi
                DataFileProgress = 100 * (o / ol)
            Loop
        ElseIf (Config(ThisJob.GameConfig).CHD = True) Then
            SourceStream.Close()
            ExtraStream.Close()
            Select Case Config(ThisJob.GameConfig).RipType
                Case "DJMAIN"
                    oi = &H1000000
                Case "FIREBEAT"
                    oi = &H100000
                Case Else
                    Exit Sub
            End Select
            sDecoderInfo = "Opening CHD file"
            Dim CHD As New clsCHD
            FileCreateFolder(ThisJob.TargetFolder)
            If CHD.chd_open_file(ThisJob.SourceFile) = clsCHD.CHD_ERROR.CHDERR_NONE Then
                If Config(ThisJob.GameConfig).RipType = "DJMAIN" Then
                    CHDStartOffset = (oi \ CHD.HunkSize)
                End If
                For x = CHDStartOffset To CHD.HunkCount - 1 Step (oi \ CHD.HunkSize)
                    'x = &H18000000 \ CHD.HunkSize
                    Select Case Config(ThisJob.GameConfig).RipType
                        Case "DJMAIN"
                            sDecoderInfo = "Extracting DJMAIN Hunk " & CStr(x)
                            DataExtractDJMAIN(CHD, CLng(x) * CLng(CHD.HunkSize))
                        Case "FIREBEAT"
                            sDecoderInfo = "Extracting FIREBEAT Hunk " & CStr(x)
                            x += ((DataExtractFIREBEAT(CHD, CLng(x) * CLng(CHD.HunkSize)) \ oi) * oi) \ CHD.HunkSize
                    End Select
                    If x < CHD.HunkCount Then
                        DataFileProgress = 100 * (x / CHD.HunkCount)
                    End If
                    If (bStopProcess) Or (bShutdown) Then
                        Exit For
                    End If
                Next
            End If
        End If
        SourceStream.Close()
        ExtraStream.Close()
        bWorking = False
    End Sub










    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        bUseSimpleMode = True
        Dim x As Integer
        Dim s As String
        If cmbSimpleGame.SelectedIndex = 0 Then
            For x = 0 To UBound(Config)
                With Config(x)
                    If Config(x).GameID <> "" Then
                        If .Info <> "" Then
                            s = sDrives(cmbSimpleCD.SelectedIndex) + .Info
                        ElseIf .File <> "" Then
                            s = sDrives(cmbSimpleCD.SelectedIndex) + .File
                        Else
                            s = ""
                        End If
                        If s <> "" Then
                            Try
                                If Dir(s) <> "" Then
                                    cmbSimpleGame.SelectedItem = Config(x).GameID
                                    TabControl1.SelectTab(tabOptions)
                                    Exit For
                                End If
                            Catch ex As Exception
                                MsgBox("The game could not be auto-detected because the media is not accessible.", MsgBoxStyle.Critical)
                                Exit Sub
                            End Try
                        End If
                    End If
                End With
            Next
            If x > UBound(Config) Then
                MsgBox("No auto-detection matches found. You may have to use the Advanced tab.", MsgBoxStyle.Information)
            End If
        Else
            TabControl1.SelectTab(tabOptions)
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        bUseSimpleMode = False
        TabControl1.SelectTab(tabOptions)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRip.Click
        QueueRip(True)
        TabControl1.SelectTab(tabQueue)
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        bShutdown = True
        My.Settings.TargetFolder = txtSimpleTarget.Text
        My.Settings.Save()
    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim x As Integer
        sDecoderInfo = "Idle"
        DebugLog("- INFO -", "Program Information", "*** BeMedia Ripper ESE" & vbCrLf & "*** Version " & Application.ProductVersion)
        PopulateDriveList()
        ConfigLoadDefinitions()
        ConfigLoadFormats()
        ConfigLoadSongDB()
        RefreshSongDB()
        SoundInit()
        For Each ConfigType In Config
            cmbAdvancedGame.Items.Add(ConfigType.Name)
            Do
                For x = 1 To cmbSimpleGame.Items.Count - 1
                    If cmbSimpleGame.Items(x) = ConfigType.GameID Then
                        Exit Do
                    End If
                Next
                If ConfigType.GameID <> "" Then
                    cmbSimpleGame.Items.Add(ConfigType.GameID)
                End If
            Loop Until True
        Next ConfigType
        cmbSimpleGame.SelectedIndex = 0
        cmbAdvancedGame.SelectedIndex = 0
        cmbBMENaming.SelectedIndex = 0
        cmbKeysoundNumbering.SelectedIndex = 0
        lblTask.Text = ""
        bProgramLoaded = True
        txtSimpleTarget.Text = My.Settings.TargetFolder
        DataInit()
        Me.Show()
    End Sub

    Private Sub PrepareSongDB()
        'this recreates the song DB for saving
        Dim Count As Integer = 1
        Dim x As Integer
        Dim y As Integer
        Dim s As String
        Dim ThisItem As DataRow
        ReDim SongDB(0)
        With SongDB(0)
            .Title = ""
            .Artist = ""
            .Genre = ""
            .SongID = 0
            .InternalName = ""
            ReDim .Difficulty(5)
            ReDim .NoteCount(5)
        End With

        For x = 0 To AutoNameTable.Rows.Count - 1
            ThisItem = AutoNameTable.Select()(x)
            ReDim Preserve SongDB(Count)
            With ThisItem
                SongDB(Count).Title = .Item("Title")
                SongDB(Count).Artist = .Item("Artist")
                SongDB(Count).Genre = .Item("Genre")
                SongDB(Count).BPM = .Item("BPM")
                SongDB(Count).InternalName = .Item("Internal Names")
                ReDim SongDB(Count).Difficulty(5)
                ReDim SongDB(Count).NoteCount(5)
                s = .Item("NoteCounts") & ","
                y = 0
                Do Until (Len(s) <= 1) Or (y = 6)
                    SongDB(Count).NoteCount(y) = Val(Strings.Left(s, InStr(s, ",") - 1))
                    s = Strings.Mid(s, InStr(s, ",") + 1)
                    y += 1
                Loop
                s = .Item("Difficulties") & ","
                y = 0
                Do Until (Len(s) <= 1) Or (y = 6)
                    SongDB(Count).Difficulty(y) = Val(Strings.Left(s, InStr(s, ",") - 1))
                    s = Strings.Mid(s, InStr(s, ",") + 1)
                    y += 1
                Loop
            End With
            Count += 1
        Next
    End Sub

    Private Sub RefreshSongDB()
        Dim NewItem As DataRow
        Dim x As Integer
        Dim s As String
        AutoNameTable = New DataTable("songDB")
        AutoNameTable.Columns.Add("Internal Names")
        AutoNameTable.Columns.Add("Title")
        AutoNameTable.Columns.Add("Artist")
        AutoNameTable.Columns.Add("Genre")
        AutoNameTable.Columns.Add("BPM")
        AutoNameTable.Columns.Add("NoteCounts")
        AutoNameTable.Columns.Add("Difficulties")
        dataAutoName.DataSource = AutoNameTable
        AutoNameTable.Clear()
        For x = 1 To UBound(SongDB)
            NewItem = AutoNameTable.NewRow
            With NewItem
                .Item("Internal Names") = SongDB(x).InternalName
                .Item("Title") = SongDB(x).Title
                .Item("Artist") = SongDB(x).Artist
                .Item("Genre") = SongDB(x).Genre
                .Item("BPM") = SongDB(x).BPM
                s = ""
                For Each nc In SongDB(x).NoteCount
                    s &= CStr(nc) & ","
                Next
                .Item("NoteCounts") = Strings.Left(s, Len(s) - 1)
                s = ""
                For Each df In SongDB(x).Difficulty
                    s &= CStr(df) & ","
                Next
                .Item("Difficulties") = Strings.Left(s, Len(s) - 1)
            End With
            AutoNameTable.Rows.Add(NewItem)
        Next
    End Sub

    Private Sub PopulateDriveList()
        Dim fso As New Scripting.FileSystemObject()
        Dim drv As Scripting.Drive
        Dim s As String
        Dim i As Integer = 1
        Dim iUseDrive As Integer = 0
        ReDim sDrives(1)
        sDrives(0) = ""
        For Each drv In fso.Drives
            s = drv.DriveLetter & ":"
            If s <> "A:" Then
                ReDim Preserve sDrives(i)
                sDrives(i) = s
                i += 1
                If drv.IsReady Then
                    If (drv.FileSystem = "CDFS") And iUseDrive = 0 Then
                        If cmbSimpleCD.SelectedIndex = -1 Then
                            iUseDrive = cmbSimpleCD.Items.Count
                        End If
                    End If
                    s += " [" & drv.VolumeName & "] (" & drv.FileSystem & ")"
                End If
                cmbSimpleCD.Items.Add(s)
            End If
        Next drv
        cmbSimpleCD.SelectedIndex = iUseDrive
    End Sub

    Private Sub SetTarget(ByVal sNewTargetFolder As String)
        txtSimpleTarget.Text = sNewTargetFolder
        txtAdvancedTarget.Text = sNewTargetFolder
        sTargetFolder = sNewTargetFolder
    End Sub

    Private Sub TargetBrowse()
        With FolderBrowser
            .ShowNewFolderButton = True
            .Description = "Browse for Target Folder"
            .RootFolder = Environment.SpecialFolder.MyComputer
            .ShowDialog()
            If .SelectedPath <> "" Then
                SetTarget(.SelectedPath)
            End If
        End With
    End Sub

    Private Sub cmdSimpleTargetBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSimpleTargetBrowse.Click
        TargetBrowse()
    End Sub

    Private Sub cmbAdvancedGame_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbAdvancedGame.SelectedIndexChanged
        If Not bProgramLoaded Then
            Exit Sub
        End If
        With Config(cmbAdvancedGame.SelectedIndex)
            txtAdvancedInterval.Text = .Interval
            txtAdvancedOffset.Text = .Offset
            If InStr(.Info, "\") Then
                lblAdvancedExtra.Text = Mid(.Info, InStrRev(.Info, "\") + 1)
            Else
                lblAdvancedExtra.Text = .Info
            End If
            chkAdvancedUseExtra.Checked = (.Info <> "")
            If cmbSimpleCD.SelectedIndex > 0 Then
                txtAdvancedSource.Text = sDrives(cmbSimpleCD.SelectedIndex) + .File
                If .Info <> "" Then
                    txtAdvancedExtra.Text = sDrives(cmbSimpleCD.SelectedIndex) + .Info
                Else
                    txtAdvancedExtra.Text = ""
                End If
            End If
        End With
        iGameType = cmbAdvancedGame.SelectedIndex
    End Sub

    Private Sub txtSimpleTarget_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSimpleTarget.TextChanged
        txtAdvancedTarget.Text = txtSimpleTarget.Text
        sTargetFolder = txtAdvancedTarget.Text
    End Sub

    Private Sub txtAdvancedTarget_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAdvancedTarget.TextChanged
        txtSimpleTarget.Text = txtAdvancedTarget.Text
    End Sub

    Private Sub cmdAdvancedSourceBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdvancedSourceBrowse.Click
        With OpenBrowser
            .CheckFileExists = True
            .CheckPathExists = True
            .Title = "Browse for Source File"
            .ShowDialog()
            If .FileName <> "" Then
                txtAdvancedSource.Text = .FileName
            End If
        End With
    End Sub

    Private Sub cmdAdvancedExtraBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdvancedExtraBrowse.Click
        With OpenBrowser
            .CheckFileExists = True
            .CheckPathExists = True
            .Title = "Browse for Info File"
            .ShowDialog()
            If .FileName <> "" Then
                txtAdvancedExtra.Text = .FileName
            End If
        End With
    End Sub

    Private Sub cmdAdvancedTargetBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdvancedTargetBrowse.Click
        TargetBrowse()
    End Sub

    Private Sub cmdDebugVAG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TabControl1.SelectTab(tabQueue)
        Rip()
    End Sub

    Private Sub MainDecoder_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles MainDecoder.RunWorkerCompleted
        bWorking = False
        sDecoderInfo = "Ready"
        SoundFileProgress = 0
    End Sub

    Private Sub tmrProgress_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrProgress.Tick
        If SoundFileProgress <= 0 Then
            prgFile.Value = 0
        ElseIf SoundFileProgress >= 100 Then
            prgFile.Value = 100
        Else
            prgFile.Value = SoundFileProgress
        End If
        If DataFileProgress <= 0 Then
            prgOperation.Value = 0
        ElseIf DataFileProgress >= 100 Then
            prgOperation.Value = 100
        Else
            prgOperation.Value = DataFileProgress
        End If
        lblDecoderStatus.Text = sDecoderInfo
    End Sub

    Private Sub txtAdvancedExtra_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtAdvancedExtra.DragDrop
        txtAdvancedExtra.Text = DragDropFiles(sender, e, 0)
    End Sub

    Private Sub txtAdvancedExtra_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtAdvancedExtra.DragEnter
        DragEnterFiles(sender, e)
    End Sub

    Private Sub txtAdvancedExtra_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAdvancedExtra.TextChanged
        sExtraFile = txtAdvancedExtra.Text
    End Sub

    Private Sub txtAdvancedSource_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtAdvancedSource.DragDrop
        txtAdvancedSource.Text = DragDropFiles(sender, e, 0)
    End Sub

    Private Sub txtAdvancedSource_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtAdvancedSource.DragEnter
        DragEnterFiles(sender, e)
    End Sub

    Private Sub txtAdvancedSource_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAdvancedSource.TextChanged
        sSourceFile = txtAdvancedSource.Text
    End Sub

    Private Sub tmrUpdate_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrUpdate.Tick
        lblChartDisplay.Text = sChartPlayback
        If bPlaying Then
            picChart.Refresh()
        End If
        If (Not bWorking) And (bQueueActive) Then
            QueueGetNextJob()
            If ThisJob.Enabled Then
                lstQueue.Items.RemoveAt(0)
                lblTask.Text = ThisJob.SourceFile
                Rip()
            Else
                bQueueActive = False
                lblTask.Text = ""
                MsgBox("All queued operations are complete.", MsgBoxStyle.Information)
                cmdQueuePerform.Text = "Begin"
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CHD As New clsCHD
        Dim hb() As Byte = {}
        'Dim f As Integer
        'Dim f2 As Integer
        'Dim x As Long
        'CHD.chd_open_file(txtAdvancedSource.Text)
        'CHD.ExtractAllHunks("c:\chd.dat")
        'f = FreeFile()
        'FileOpen(f, "c:\chdexp.dat", OpenMode.Binary)
        'f2 = FreeFile()
        'FileOpen(f2, "c:\chdexpBSW.dat", OpenMode.Binary)
        'For x = 0 To CHD.HunkCount - 1 Step (&H1000000 \ CHD.HunkSize)
        '    CHD.ReadHunkBytes(x * CLng(CHD.HunkSize), CHD.HunkSize, hb)
        '    FilePut(f, hb)
        '    DataByteSwap(hb)
        '    FilePut(f2, hb)
        'Next
        'FileClose(f)
        'FileClose(f2)
        'CHD.ExtractAllValidHunks("c:\chd.raw")
    End Sub

    Private Sub cmdPlayback_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPlayback.Click
        ChartFile = TextBox22.Text
        If cmdPlayback.Text = "Play" Then
            WindowHandle = Me.Handle
            PlaybackThread = New Thread(New ThreadStart(AddressOf Playback))
            PlaybackThread.Start()
            cmdPlayback.Text = "Stop"
        ElseIf cmdPlayback.Text = "Stop" Then
            cmdPlayback.Text = "Wait..."
            If bPlaying Then
                bStopPlaying = True
                Do While PlaybackThread.IsAlive
                    Application.DoEvents()
                Loop
            End If
            cmdPlayback.Text = "Play"
        End If

    End Sub

    Private Sub Playback()
        Thread.CurrentThread.Priority = ThreadPriority.AboveNormal
        Dim x As Integer
        Dim ChartFolder As String = ""
        Dim ThisBPM As Double
        Dim LastTime As Integer
        Dim InitialMetric As Double
        Dim InitialTime As Integer
        Dim DSound As clsDirectSound = New clsDirectSound
        If InStr(ChartFile, "\") > 0 Then
            ChartFolder = Strings.Left(ChartFile, InStrRev(ChartFile, "\"))
        End If
        If Not DSound.IsReady Then
            If Not DSound.Init(WindowHandle) Then
                MsgBox("There was a problem initializing DirectSound. You won't be able to use chart playback.", MsgBoxStyle.Information)
                DSound = Nothing
                Exit Sub
            End If
        End If
        ChartLoadFile(ChartFile, ChartTypes.BME)
        For x = 1 To UBound(Chart.WaveTable)
            If Chart.WaveTable(x) <> "" Then
                DSound.SoundLoad(ChartFolder & Chart.WaveTable(x), x)
            End If
        Next
        ThisTime = 0
        ThisEvent = 0
        ThisMetric = 0
        LastTime = 0
        bPlaying = True
        MainTimer.Start()
        Do While (ThisEvent < Chart.EventCount) And (bPlaying) And (Not bStopPlaying) And (Not bShutdown)
            MainTimer.Update()
            sChartPlayback = CStr(Int(MainTimer.Duration * 1000)) & "(" & CStr(ThisEvent) & "/" & CStr(Chart.EventCount) & ")"
            LastTime = ThisTime
            ThisTime = CInt(MainTimer.Duration * 1000)
            ThisMetric = InitialMetric + ((ThisBPM / 4) / 60000) * (ThisTime - InitialTime)
            Do While ThisTime >= Chart.Events(ThisEvent).OffsetMSec
                With Chart.Events(ThisEvent)
                    If .Parameter = ChartParameter.Note Or .Parameter = ChartParameter.BGM Then
                        DSound.SoundPlay(.Value)
                    ElseIf .Parameter = ChartParameter.TempoChange Then
                        ThisBPM = .Value
                        InitialTime = .OffsetMSec
                        InitialMetric = .OffsetMetric
                    ElseIf .Parameter = ChartParameter.EndSong Then
                        bPlaying = False
                        Exit Do
                    End If
                End With
                ThisEvent += 1
                ThisMetric = Chart.Events(ThisEvent).OffsetMetric
            Loop
            Application.DoEvents()
        Loop
        MainTimer.Stop()
        DSound.SoundFreeAll()
        bStopPlaying = False
    End Sub

    Private Sub picChart_Paint1(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles picChart.Paint
        'Dim x As Integer
        Dim y As Integer = ThisEvent
        Dim z As Integer
        Dim a As Integer
        Dim ThisColor As Brush = Brushes.Black
        If bPlaying Then
            e.Graphics.Clear(Color.Black)
            e.Graphics.DrawLine(Pens.White, 128, 0, 128, picChart.Height)
            Do While y < Chart.EventCount
                With Chart.Events(y)
                    z = ((picChart.Height - ((.OffsetMetric - ThisMetric) * picChart.Height * 0.8))) - 4
                    If .Parameter = ChartParameter.Note Then
                        If .Lane <= 7 Then
                            If .Player = 1 Then
                                a = (16 * .Lane)
                            ElseIf .Player = 2 Then
                                a = (16 * 8) + (16 * .Lane)
                            End If
                            If z < 0 Then
                                Exit Do
                            End If
                            Select Case Chart.ChartType
                                Case ChartTypes.IIDXAC, ChartTypes.IIDXCS, ChartTypes.IIDXCS2, ChartTypes.IIDXCS5, ChartTypes.BME
                                    Select Case .Lane
                                        Case 0, 2, 4, 6
                                            ThisColor = Brushes.LightGray
                                        Case 1, 3, 5
                                            ThisColor = Brushes.Blue
                                        Case 7
                                            ThisColor = Brushes.Red
                                    End Select
                            End Select
                            e.Graphics.FillRectangle(ThisColor, New Rectangle(a, z - 4, 15, 4))

                        End If
                    ElseIf .Parameter = ChartParameter.Measure Then
                        e.Graphics.DrawLine(Pens.White, 0, z, picChart.Width, z)
                    End If
                End With
                y += 1
            Loop
        End If
    End Sub

    Private Sub TextBox22_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox22.DragDrop
        TextBox22.Text = DragDropFiles(sender, e, 0)
    End Sub

    Private Sub TextBox22_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox22.DragEnter
        DragEnterFiles(sender, e)
    End Sub

    Private Sub cmdQueuePerform_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQueuePerform.Click
        If cmdQueuePerform.Text = "Stop" Then
            cmdQueuePerform.Text = "Begin"
            bQueueActive = False
        Else
            cmdQueuePerform.Text = "Stop"
            bQueueActive = True
        End If
    End Sub

    Private Sub cmdQueue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQueue.Click
        QueueRip(False)
        TabControl1.SelectTab(tabQueue)
    End Sub

    Private Sub cmdQueueRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQueueRemove.Click
    End Sub

    Private Sub cmdQueueClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQueueClear.Click
        lstQueue.Items.Clear()
        QueueClear()
    End Sub

    Private Sub cmdAutoNameSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAutoNameSave.Click
        PrepareSongDB()
        ConfigSaveSongDB()
    End Sub

    Private Sub cmdAutoNameReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAutoNameReset.Click
        ConfigLoadSongDB()
        RefreshSongDB()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim f() As Byte = My.Computer.FileSystem.ReadAllBytes("h:\thing\06C ZERO.raw")
        Dim g() As Byte = My.Computer.FileSystem.ReadAllBytes("h:\thing\06U ZERO.raw")
        Dim h() As Byte = My.Computer.FileSystem.ReadAllBytes("h:\thing\07C ZERO.raw")
        Dim i() As Byte = My.Computer.FileSystem.ReadAllBytes("h:\thing\07U ZERO.raw")
        DataDumpArrayXOR("h:\thing\ZERO 6C-6U.raw", f, g)
        DataDumpArrayXOR("h:\thing\ZERO 6C-7C.raw", f, h)
        DataDumpArrayXOR("h:\thing\ZERO 6U-7U.raw", g, i)
    End Sub
End Class
