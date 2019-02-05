Option Explicit On
Module modConfig
    Public Structure ConfigType
        Public Name As String
        Public File As String
        Public Info As String
        Public TableOffset As Long
        Public SongOffset As Long
        Public TableType As Integer
        Public SongType As Integer
        Public Index As Integer
        Public Outdex As Integer
        Public ReadEXE As Boolean
        Public Offset As Long
        Public Interval As Long
        Public CDName As String
        Public UseKey As String
        Public RipType As String
        Public Titles As Integer
        Public CHD As Boolean
        Public GameID As String
        Public DataID As Integer
    End Structure

    Public Structure ConfigKeyType
        Public Name As String
        Public Signature As String
        Public Block As String
        Public DecodeType As String
        Public Function AsBytes() As Byte()
            Dim r() As Byte = ConfigConvertKeyToBytes(Block)
            Return r
        End Function
    End Structure

    Public Structure ConfigFormatType
        Public Format As Integer
        Public Size As Integer
        Public NameO As Integer
        Public NameT As Integer
        Public NameS As Integer
        Public NameL As Integer
        Public NameP As Integer
        Public NameSZ As Integer
        Public DiffC As Integer
        Public DiffT As Integer
        Public DiffO As Integer
        Public DiffS As String
        Public DiffS5 As String
        Public DiffL As String
        Public KeyC As Integer
        Public KeyT As Integer
        Public KeyO As Integer
        Public KeyP As Integer
        Public KeyS As String
        Public SetC As Integer
        Public SetT As Integer
        Public SetO As Integer
        Public SetS As String
        Public Key5 As Integer
        Public MovieC As Integer
        Public MovieT As Integer
        Public MovieO As Integer
        Public MovieD As Integer
        Public ChartAdjust As Integer
        Public Timing As Integer
        Public DatFile As Integer
    End Structure

    Public Structure ConfigFileSystemType
        Public FileTbl As Integer
        Public Size As Integer
        Public Offset As Integer
        Public Length As Integer
        Public OffMult As Integer
        Public LenMult As Integer
        Public BackChk As Integer
        Public Count As Integer
    End Structure

    Public Structure ConfigSongDBType
        Public Title As String
        Public Artist As String
        Public Genre As String
        Public SongID As Integer
        Public Difficulty() As Integer
        Public NoteCount() As Integer
        Public InternalName As String
        Public BPM As Integer
    End Structure

    Public Structure RipFormatType
        Public Index As Integer
        Public Size As Integer
        Public NameType As Integer
        Public NameShort As Integer
        Public NameLong As Integer
        Public NameSize As Integer
        Public Name As String
        Public DiffCount As Integer
        Public DiffType As Integer
        Public Diff() As Integer
        Public KeyCount As Integer
        Public KeyType As Integer
        Public Key() As Integer
        Public SetsCount As Integer
        Public SetsType As Integer
        Public Sets() As Integer
        Public MovieCount As Integer
        Public MovieType As Integer
        Public Movie As Integer
    End Structure


    Public Config() As ConfigType
    Public SongDB() As ConfigSongDBType
    Public DecKeys() As ConfigKeyType
    Public Formats() As ConfigFormatType
    Public FileSystems() As ConfigFileSystemType

    Public Function ConfigConvertKeyToBytes(ByVal sKey As String) As Byte()
        Dim R As Byte() = {}
        Dim x As Integer
        If Len(sKey) > 1 Then
            ReDim R(0 To (Len(sKey) \ 2) - 1)
            For x = 1 To Len(sKey) Step 2
                R((x - 1) \ 2) = Val("&H" & Strings.Mid(sKey, x, 2))
            Next
        End If
        Return R
    End Function


    Public Sub ConfigLoadDefinitions()
        Dim b() As Byte
        Dim o As Integer = 0
        Dim s As String
        Dim v As String
        Dim a As Integer = 0
        Dim x As Integer = 0
        Dim iIndex As Integer = 1
        Dim iKeyIndex As Integer = 0
        Dim bBeginEnabled As Boolean = False
        Dim bKeyEnabled As Boolean = False
        ReDim b(0)
        ReDim Config(0 To 1)
        ReDim DecKeys(0)
        On Error Resume Next
        With Config(0) 'default, this one just means read generic formats in an unknown file
            .File = ""
            .GameID = ""
            .Index = 0
            .Info = ""
            .Interval = &H800&
            .Name = "[none]"
            .Offset = 0
            .Outdex = 0
            .ReadEXE = False
            .SongOffset = 0
            .SongType = 0
            .TableOffset = 0
            .TableType = 0
            .DataID = -1
        End With
        If FileExists(FileAppPath() + "BeMediaDefs.ini") Then
            FileLoadMemory(b, FileAppPath() + "BeMediaDefs.ini")
            Do While o < UBound(b)
                s = Trim(FileGetLine(b, o))
                a = InStr(s, "//")
                If a > 0 Then
                    s = Left(s, a - 1)
                End If
                a = InStr(s, Chr(34))
                If a > 0 Then
                    v = UCase(Trim(Left(s, a - 1)))
                    s = Mid(s, a + 1)
                    a = InStr(s, Chr(34))
                    If a > 0 Then
                        s = Left(s, a - 1)
                    End If
                Else
                    v = Trim(s)
                End If
                If (bBeginEnabled = True) Or (bBeginEnabled = False And v = "BEGIN") Then
                    If s <> "" Then
                        Select Case v
                            Case "BEGIN"
                                bBeginEnabled = True
                                ReDim Preserve Config(0 To iIndex)
                                With Config(iIndex) 'set defaults
                                    .Name = s
                                    .File = ""
                                    .Info = ""
                                    .SongOffset = 0
                                    .SongType = -1
                                    .TableOffset = 0
                                    .TableType = -1
                                    .Index = 0
                                    .Outdex = 0
                                    .ReadEXE = False
                                    .Offset = 0
                                    .Interval = &H800&
                                    .RipType = ""
                                    .UseKey = ""
                                    .CDName = ""
                                    .Titles = -1
                                    .DataID = -1
                                End With
                            Case "GAMEID"
                                Config(iIndex).GameID = s
                            Case "FILE"
                                Config(iIndex).File = s
                            Case "INFO"
                                Config(iIndex).Info = s
                            Case "TABLE"
                                Config(iIndex).TableOffset = CLng(s)
                            Case "SONG"
                                Config(iIndex).SongOffset = CLng(s)
                            Case "OFFSET"
                                Config(iIndex).Offset = CLng(s)
                            Case "INTERVAL"
                                Config(iIndex).Interval = CLng(s)
                            Case "TABLEID"
                                Config(iIndex).TableType = CInt(s)
                            Case "SONGID"
                                Config(iIndex).SongType = CInt(s)
                            Case "INDEX"
                                Config(iIndex).Index = CInt(s)
                            Case "OUTDEX"
                                Config(iIndex).Outdex = CInt(s)
                            Case "READEXE"
                                Config(iIndex).ReadEXE = (s = "1")
                            Case "CHD"
                                Config(iIndex).CHD = (s = "1")
                            Case "RIPTYPE"
                                Config(iIndex).RipType = s
                            Case "USEKEY"
                                Config(iIndex).UseKey = s
                            Case "CDNAME"
                                Config(iIndex).CDName = s
                            Case "TITLES"
                                Config(iIndex).Titles = s
                            Case "DATAID"
                                Config(iIndex).DataID = CInt(s)
                            Case "END"
                                bBeginEnabled = False
                                iIndex += 1
                            Case Else
                        End Select
                    End If
                ElseIf (bKeyEnabled = True) Or (bKeyEnabled = False And v = "KEY") Then
                    Select Case v
                        Case "KEY"
                            ReDim Preserve DecKeys(iKeyIndex)
                            With DecKeys(iKeyIndex)
                                .Name = v
                                .Block = ""
                                .DecodeType = ""
                                .Signature = ""
                            End With
                            bKeyEnabled = True
                        Case "SIG"
                            DecKeys(iKeyIndex).Signature = v
                        Case "SIGHEX"
                            DecKeys(iKeyIndex).Signature = ""
                            For x = 1 To Len(v) Step 2
                                DecKeys(iKeyIndex).Signature += Chr(Val("&H" + Mid(v, x, 2)))
                            Next x
                        Case "DEC"
                            DecKeys(iKeyIndex).DecodeType = v
                        Case "BLOCK"
                            DecKeys(iKeyIndex).Block = v
                        Case "END"
                            bKeyEnabled = False
                            iKeyIndex += 1
                    End Select
                End If
            Loop
        End If
        DebugLog("ConfigLoadDefinitions", "Definitions Loaded Successfully", "Files:" + CStr(iIndex) + vbCrLf + "Keys:" + CStr(iKeyIndex))
    End Sub

    Public Sub ConfigLoadFormats()
        Dim b() As Byte
        Dim o As Integer = 0
        Dim s As String
        Dim v As String
        Dim a As Integer = 0
        Dim x As Integer = 0
        Dim iIndex As Integer = 0
        Dim iSystemIndex As Integer = 0
        Dim bBeginEnabled As Boolean = False
        Dim bSystemEnabled As Boolean = False
        ReDim b(0)
        ReDim Formats(0)
        ReDim FileSystems(0)
        On Error Resume Next
        If FileExists(FileAppPath() + "BeMediaInfoFormats.ini") Then
            FileLoadMemory(b, FileAppPath() + "BeMediaInfoFormats.ini")
            Do While o < UBound(b)
                s = Trim(FileGetLine(b, o))
                a = InStr(s, "//")
                If a > 0 Then
                    s = Left(s, a - 1)
                End If
                a = InStr(s, Chr(34))
                If a > 0 Then
                    v = UCase(Trim(Left(s, a - 1)))
                    s = Mid(s, a + 1)
                    a = InStr(s, Chr(34))
                    If a > 0 Then
                        s = Left(s, a - 1)
                    End If
                Else
                    v = Trim(s)
                End If
                If (bBeginEnabled = True) Or (bBeginEnabled = False And v = "FORMAT") Then
                    If s <> "" Then
                        Select Case v
                            Case "FORMAT"
                                bBeginEnabled = True
                                ReDim Preserve Formats(0 To iIndex)
                                With Formats(iIndex) 'set defaults
                                    .Format = CInt(s)
                                    .Size = 0
                                    .NameT = 0
                                    .NameS = 0
                                    .NameL = 0
                                    .NameP = 0
                                    .NameSZ = 0
                                    .NameO = -1
                                    .DiffC = 0
                                    .DiffT = 0
                                    .DiffO = -1
                                    .DiffS = ""
                                    .DiffL = ""
                                    .KeyC = 0
                                    .KeyT = 0
                                    .KeyO = -1
                                    .KeyP = 0
                                    .SetC = 0
                                    .SetT = 0
                                    .SetO = -1
                                    .SetS = ""
                                    .MovieC = 0
                                    .MovieT = 0
                                    .MovieO = -1
                                    .DatFile = -1
                                End With
                            Case "SIZE"    'size of the structure in bytes
                                Formats(iIndex).Size = CInt(s)
                            Case "NAMET"   'type of name, 0=string or 4=pointer
                                Formats(iIndex).NameT = CInt(s)
                            Case "NAMES"   'short name pointer
                                Formats(iIndex).NameS = CInt(s)
                            Case "NAMEL"   'long name pointer
                                Formats(iIndex).NameL = CInt(s)
                            Case "NAMEP"   'pointer offset
                                Formats(iIndex).NameP = CInt(s)
                            Case "NAMESZ"  'non-pointer string length
                                Formats(iIndex).NameSZ = CInt(s)
                            Case "NAME"    'name offset
                                Formats(iIndex).NameO = CInt(s)
                            Case "DIFFC"   'difficulty count
                                Formats(iIndex).DiffC = CInt(s)
                            Case "DIFFT"   'difficulty type
                                Formats(iIndex).DiffT = CInt(s)
                            Case "DIFF"    'difficulty offset
                                Formats(iIndex).DiffO = CInt(s)
                            Case "DIFFS"   'difficulty strings
                                Formats(iIndex).DiffS = s
                            Case "DIFFL"   'difficulty indexes
                                Formats(iIndex).DiffL = s
                            Case "KEYS"    'chart string
                                Formats(iIndex).KeyS = s
                            Case "KEYC"    'chart count
                                Formats(iIndex).KeyC = CInt(s)
                            Case "KEYT"    'chart type
                                Formats(iIndex).KeyT = CInt(s)
                            Case "KEY"     'chart table offset
                                Formats(iIndex).KeyO = CInt(s)
                            Case "KEYP"    'chart pointer
                                Formats(iIndex).KeyP = CInt(s)
                            Case "SETC"    'key/bgm count
                                Formats(iIndex).SetC = CInt(s)
                            Case "SETT"    'key/bgm type
                                Formats(iIndex).SetT = CInt(s)
                            Case "SET"     'key/bgm table offset
                                Formats(iIndex).SetO = CInt(s)
                            Case "SETS"    'key/bgm strings
                                Formats(iIndex).SetS = s
                            Case "MOVIEC"  'movie count
                                Formats(iIndex).MovieC = CInt(s)
                            Case "MOVIET"  'movie type
                                Formats(iIndex).MovieT = CInt(s)
                            Case "MOVIE"   'movie table offset
                                Formats(iIndex).MovieO = CInt(s)
                            Case "MOVIED"  'movie dupe value
                                Formats(iIndex).MovieD = CInt(s)
                            Case "5KEY"    '5-key switch, 0=5k 1=7k (for BMUS)
                                Formats(iIndex).Key5 = CInt(s)
                            Case "5DIFFS"
                                Formats(iIndex).DiffS5 = s
                            Case "ADJUST"
                                Formats(iIndex).ChartAdjust = CInt(s)
                            Case "TIMING"
                                Formats(iIndex).Timing = CInt(s)
                            Case "DATFILE"
                                Formats(iIndex).DatFile = CInt(s)
                            Case "END"     'end of definition
                                bBeginEnabled = False
                                iIndex += 1
                            Case Else
                        End Select
                    End If
                ElseIf (bSystemEnabled = True) Or (bSystemEnabled = False And v = "FILETBL") Then
                    Select Case v
                        Case "FILETBL"
                            ReDim Preserve FileSystems(iSystemIndex)
                            With FileSystems(iSystemIndex)
                                .FileTbl = CInt(s)
                                .Size = 0
                                .Offset = 0
                                .Length = 0
                                .OffMult = 1
                                .LenMult = 1
                                .BackChk = 0
                                .Count = 0
                            End With
                            bSystemEnabled = True
                        Case "SIZE"
                            FileSystems(iSystemIndex).Size = CInt(s)
                        Case "OFFSET"
                            FileSystems(iSystemIndex).Offset = CInt(s)
                        Case "LENGTH"
                            FileSystems(iSystemIndex).Length = CInt(s)
                        Case "OFFMULT"
                            FileSystems(iSystemIndex).OffMult = CInt(s)
                        Case "LENMULT"
                            FileSystems(iSystemIndex).LenMult = CInt(s)
                        Case "BACKCHK"
                            FileSystems(iSystemIndex).BackChk = CInt(s)
                        Case "COUNT"
                            FileSystems(iSystemIndex).Count = CInt(s)
                        Case "END"
                            bSystemEnabled = False
                            iSystemIndex += 1
                        Case Else
                    End Select
                End If
            Loop
        End If
        DebugLog("ConfigLoadFormats", "Formats Loaded Successfully", "Formats:" + CStr(iIndex) + vbCrLf + "Filesystems:" + CStr(iSystemIndex))
    End Sub

    'retrieve by title
    Public Function ConfigGetSongDBInfo(ByVal SongTitle As String) As ConfigSongDBType
        Dim ReturnDB As ConfigSongDBType
        With ReturnDB
            .Title = ""
            .Artist = ""
            .Genre = ""
            .InternalName = ""
            .SongID = 0
            ReDim .Difficulty(5)
            ReDim .NoteCount(5)
        End With
        Dim x As Integer
        For x = 0 To UBound(SongDB)
            With SongDB(x)
                If UCase(.Title) = UCase(SongTitle) Then
                    ReturnDB = SongDB(x)
                    Exit For
                End If
            End With
        Next
        Return ReturnDB
    End Function

    'retrieve by multiple note counts and one bpm
    Public Function ConfigGetSongDBInfo(ByVal NoteCounts() As Integer, ByVal BPM As Integer) As ConfigSongDBType
        Dim ReturnDB As ConfigSongDBType
        With ReturnDB
            .Title = ""
            .Artist = ""
            .Genre = ""
            .InternalName = ""
            .SongID = 0
            ReDim .Difficulty(5)
            ReDim .NoteCount(5)
        End With
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim e As Boolean
        Dim f As Boolean
        If BPM = 0 Or NoteCounts(0) = 0 Then
            Return SongDB(0)
            Exit Function
        End If
        For x = 0 To UBound(SongDB)
            With SongDB(x)
                For y = 0 To 5
                    If .Difficulty(y) <> 0 Then
                        Exit For
                    End If
                Next
                If y = 6 Then
                    If BPM = .BPM Then
                        f = False
                        For z = 0 To UBound(NoteCounts)
                            If NoteCounts(z) > 0 Then
                                e = True
                                For y = 0 To 5
                                    If .NoteCount(y) > 0 Then
                                        If .NoteCount(y) = NoteCounts(z) Then
                                            e = False
                                        End If
                                    End If
                                Next
                                f = (f Or e)
                            End If
                        Next
                        If Not f Then
                            Return SongDB(x)
                            Exit Function
                        End If
                    End If
                End If
            End With
        Next
        Return ReturnDB
    End Function

    'retrieve by note count and bpm
    Public Function ConfigGetSongDBInfo(ByVal NoteCount As Integer, ByVal BPM As Integer) As ConfigSongDBType
        Dim ReturnDB As ConfigSongDBType
        With ReturnDB
            .Title = ""
            .Artist = ""
            .Genre = ""
            .InternalName = ""
            .SongID = 0
            ReDim .Difficulty(5)
            ReDim .NoteCount(5)
        End With
        If NoteCount > 0 Then
            Dim x As Integer
            Dim y As Integer
            For x = 0 To UBound(SongDB)
                With SongDB(x)
                    For y = 0 To 5
                        If .Difficulty(y) <> 0 Then
                            Exit For
                        End If
                    Next
                    If y = 6 Then
                        If BPM = .BPM Then
                            For y = 0 To 5
                                If (.NoteCount(y) > 0) AndAlso (NoteCount = .NoteCount(y)) Then
                                    Return SongDB(x)
                                    Exit Function
                                End If
                            Next
                        End If
                    End If
                End With
            Next
        End If
        Return ReturnDB
    End Function

    Public Sub ConfigSaveSongDB()
        Dim Writer As New IO.StreamWriter(FileAppPath() + "SongDB.ini", False)
        Dim x As Integer
        Dim y As Integer
        For x = 1 To UBound(SongDB)
            Dim s As String
            With SongDB(x)
                s = .Title
                If .InternalName <> "" Then
                    s &= "|" & .InternalName
                End If
                s &= vbTab
                s &= .Artist & vbTab
                s &= .Genre & vbTab
                s &= CStr(.BPM) & vbTab
                For y = 0 To 5
                    s &= CStr(.Difficulty(y)) & vbTab
                Next
                For y = 0 To 5
                    s &= CStr(.NoteCount(y)) & vbTab
                Next
                Writer.WriteLine(s)
            End With
        Next
        Writer.Close()
    End Sub


    Public Sub ConfigLoadSongDB()
        Dim s As String
        Dim a As Integer
        Dim o As Integer
        Dim n As Integer
        Dim v As String
        Dim b() As Byte = {}
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
        Dim SongDBCount As Integer = 1
        If FileExists(FileAppPath() + "SongDB.ini") Then
            FileLoadMemory(b, FileAppPath() + "SongDB.ini")
            Do While o < UBound(b)
                s = Trim(FileGetLine(b, o))
                a = InStr(s, vbTab)
                n = 0
                If a > 0 Then
                    s &= vbTab
                    ReDim Preserve SongDB(SongDBCount)
                    SongDB(SongDBCount) = New ConfigSongDBType
                    ReDim SongDB(SongDBCount).Difficulty(5)
                    ReDim SongDB(SongDBCount).NoteCount(5)
                    Do While a > 0
                        v = Left(s, a - 1)
                        s = Mid(s, a + 1)
                        With SongDB(SongDBCount)
                            Select Case n
                                Case 0 'title
                                    If InStr(v, "|") > 0 Then
                                        .InternalName = Mid(v, InStr(v, "|") + 1)
                                        .Title = Left(v, InStr(v, "|") - 1)
                                    Else
                                        .Title = v
                                        .InternalName = v
                                    End If
                                Case 1 'artist
                                    .Artist = v
                                Case 2 'genre
                                    .Genre = v
                                Case 3 'bpm
                                    .BPM = Val(v)
                                Case 4, 5, 6, 7, 8, 9
                                    .Difficulty(n - 4) = Val(v)
                                Case 10, 11, 12, 13, 14, 15
                                    .NoteCount(n - 10) = Val(v)
                            End Select
                        End With
                        a = InStr(s, vbTab)
                        n += 1
                    Loop
                    SongDBCount += 1
                End If
            Loop
        End If
    End Sub

End Module
