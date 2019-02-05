Option Explicit On
Module modSound
    Private Const MaxReadBuffer = 8192
    Private VAGCoEff(5, 2) As Integer

    'these are all the known frequencies - unknown ones are interpolated linearly between these
    'when making changes, keep the second column in order
    Private FreqTable() As Integer = { _
        44100, 15940, _
        41624, 16196, _
        40000, 16491, _
        36000, 16642, _
        37000, 16703, _
        37000, 16708, _
        37818, 16762, _
        36000, 16899, _
        35002, 16964, _
        33600, 17222, _
        32000, 17476, _
        32000, 17533, _
        30000, 17774, _
        28000, 18005, _
        24000, 18432, _
        22050, 19007, _
        22050, 19012, _
        11025, 19714, _
        16000, 20605}

    Private Structure SoundDecodeBuffer
        Public Data() As Short
    End Structure


    Public Enum SoundDecodeType
        None
        IIDX3                   'IIDX 3rdstyle PS2
        IIDX9                   'IIDX 9thstyle PS2
        IIDX11Keysound          'IIDX 11thstyle PS2
        IIDX11BGM
        IIDX14Keysound          'IIDX 14thstyle PS2
        IIDX14BGM
        BMDXgeneric             'test...
        DDRPSX                  'playstation DDR
        DSDISNEY                'playstation Dancing Stage
        DDRPS2                  'playstation2 DDR
        DDRSN                   'playstation2 DDR (supernova+)
        DJMAIN                  'arcade beatmania
        FIREBEAT                'arcade beatmaniaIIDX
    End Enum

    Public Structure SoundRipInstance
        Public OutputName As String
        Public FileOffset As Long
        Public Freq As Integer
        Public BlockSize As Integer
        Public Channels As Integer
        Public LoopStart As Integer
        Public LoopEnd As Integer
        Public VolumeL As Single
        Public VolumeR As Single
        Public Pan As Single
        Public DecodeType As SoundDecodeType
        Public DecodeParam As Integer
        Public MaxLength As Integer
    End Structure

    Public Structure SoundRipParameters
        Public DoRip As Boolean
        'Public FileNumber As Integer
        Public FileStream As IO.Stream
        Public InstanceCount As Integer
        Public Instances() As SoundRipInstance
    End Structure

    Public SoundRipInfo As SoundRipParameters

    Public Sub SoundInit()
        VAGCoEff(0, 0) = 0
        VAGCoEff(0, 1) = 0
        VAGCoEff(1, 0) = 60
        VAGCoEff(1, 1) = 0
        VAGCoEff(2, 0) = 115
        VAGCoEff(2, 1) = -52
        VAGCoEff(3, 0) = 98
        VAGCoEff(3, 1) = -55
        VAGCoEff(4, 0) = 122
        VAGCoEff(4, 1) = -60
        With SoundRipInfo
            .DoRip = False
            .FileStream = IO.Stream.Null
            .InstanceCount = 0
            ReDim .Instances(0)
        End With
    End Sub

    Public Sub SoundStartDecode(ByVal sSourceFile As String, ByVal sTargetPath As String, Optional ByVal sPrefix As String = "")
        Dim x As Integer
        Dim s As String
        Dim str As IO.Stream
        sTargetPath = Trim(sTargetPath)
        s = sTargetPath
        SoundFileProgress = 0
        If sPrefix <> "" Then
            x = x
        End If
        If (Strings.Right(sTargetPath, 1) = "\") Then
            Try
                MkDir(sTargetPath)
            Catch ex As Exception
                DebugLog("SoundStartDecode", "Could not create folder for multiple sounds in this set.")
            End Try
        End If
        With SoundRipInfo
            If .DoRip = False Then
                Exit Sub
            End If
            str = New IO.FileStream(sSourceFile, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.ReadWrite)
            x = 0
            For Each SoundRipInstance In SoundRipInfo.Instances
                With SoundRipInstance
                    Select Case .DecodeType
                        Case SoundDecodeType.None
                        Case SoundDecodeType.DJMAIN
                        Case SoundDecodeType.FIREBEAT
                        Case Else
                            SoundDecodeVAGStream(sTargetPath & sPrefix & Trim(SoundRipInstance.OutputName) & ".wav", str, .Channels, .BlockSize, .FileOffset, .Freq, .DecodeType, .DecodeParam, .VolumeL, .VolumeR, .MaxLength)
                    End Select

                End With
                x += 1
                If SoundRipInfo.InstanceCount > 1 Then
                    SoundFileProgress = (CSng(x) / SoundRipInfo.InstanceCount) * 100
                End If
            Next
        End With
        SoundFileProgress = 0
        str.Close()
    End Sub

    Public Sub SoundWriteWAVHeader(ByVal FileNumber As Integer, ByVal Freq As Integer, ByVal Channels As Short)
        Dim x As UInteger
        x = &H46464952 : FilePut(FileNumber, x, 1)   ' "RIFF"
        x = LOF(FileNumber) - 8 : FilePut(FileNumber, x, 5) 'riff filesize
        x = &H45564157 : FilePut(FileNumber, x, 9)   ' "WAVE"
        x = &H20746D66 : FilePut(FileNumber, x, 13)  ' "fmt"
        x = &H10 : FilePut(FileNumber, x, 17)        ' headersize
        x = (&H10000 * Channels) + 1 : FilePut(FileNumber, x, 21) ' format + channels
        FilePut(FileNumber, Freq, 25)                ' frequency
        Freq = Freq * Channels * 2 : FilePut(FileNumber, Freq, 29) ' bytes per second
        x = &H100004 : FilePut(FileNumber, x, 33) ' 16 bits, 4 bytes/sample
        x = &H61746164 : FilePut(FileNumber, x, 37)  ' "data"
        x = LOF(FileNumber) - 44 : FilePut(FileNumber, x, 41) 'data filesize
    End Sub

    Public Function SoundWAVHeaderSize() As Integer
        SoundWAVHeaderSize = 44
    End Function

    Public Sub SoundUpsampleSave(ByVal xSource() As Byte, ByVal FileName As String, ByVal bFlipSign As Boolean, ByVal iBits As Integer, ByVal iChannels As Integer, ByVal iFreq As Integer, ByVal VolL As Double, ByVal VolR As Double, Optional ByVal RemoveSilence As Boolean = True)
        Dim outp() As Short = {}
        SoundUpsample(xSource, outp, bFlipSign, iBits, iChannels, iFreq, VolL, VolR)
        SoundRemoveSilence(outp)
        'Dim f As Integer = FreeFile()
        'FileOpen(f, FileName, OpenMode.Binary, OpenAccess.Write, OpenShare.Default)
        'FilePut(f, outp, SoundWAVHeaderSize() + 1)
        'SoundWriteWAVHeader(f, iFreq, 2)
        'FileClose(f)
        SoundSave(outp, FileName, iFreq)
    End Sub

    Public Sub SoundSave(ByVal xSource() As Short, ByVal FileName As String, ByVal iFreq As Integer)
        Dim f As Integer = FreeFile()
        Dim x As Integer
        Dim y As Integer = 0
        Dim FO As IO.FileStream = New IO.FileStream(FileName, IO.FileMode.Create)
        Dim bo() As Byte
        ReDim bo((UBound(xSource) * 2) + 1)
        For x = 0 To UBound(xSource)
            bo(y) = xSource(x) And 255
            y += 1
            bo(y) = (xSource(x) >> 8) And 255
            y += 1
        Next
        FO.Position = SoundWAVHeaderSize()
        FO.Write(bo, 0, UBound(bo) + 1)
        FO.Close()
        FO = Nothing
        FileOpen(f, FileName, OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
        SoundWriteWAVHeader(f, iFreq, 2)
        FileClose(f)
    End Sub

    Public Sub SoundRemoveSilence(ByRef xSound() As Short, Optional ByVal bTrailing As Boolean = True, Optional ByVal bLeading As Boolean = False)
        Dim FirstNonSilence As Integer = -1
        Dim LastNonSilence As Integer = -1
        Dim x As Integer
        For x = 0 To UBound(xSound) - 2 Step 2
            If xSound(x) <> 0 Or xSound(x + 1) <> 0 Then
                LastNonSilence = x
                If (FirstNonSilence <> -1) Then
                    FirstNonSilence = x
                End If
            End If
        Next
        If bTrailing Then
            If LastNonSilence > -1 Then
                ReDim Preserve xSound(0 To LastNonSilence + 1)
            End If
        End If
    End Sub

    Public Sub SoundCombineSave(ByVal xLeft() As Short, ByVal xRight() As Short, ByVal FileName As String, ByVal iFreq As Integer)
        Dim SaveStream As IO.FileStream
        Dim Writer As IO.BinaryWriter
        If UBound(xLeft) > UBound(xRight) Then
            ReDim Preserve xRight(UBound(xLeft))
        ElseIf UBound(xLeft) < UBound(xRight) Then
            ReDim Preserve xLeft(UBound(xRight))
        End If
        Dim outp() As Short
        Dim x As Integer
        ReDim outp(UBound(xLeft))
        For x = 0 To UBound(outp) Step 2
            outp(x) = xLeft(x)
            outp(x + 1) = xRight(x + 1)
        Next
        SaveStream = New IO.FileStream(FileName, IO.FileMode.Create)
        Writer = New IO.BinaryWriter(SaveStream)
        Dim f As Integer = FreeFile()
        Try
            SaveStream.Position = SoundWAVHeaderSize()
            For x = 0 To UBound(outp)
                Writer.Write(outp(x))
            Next
            SaveStream.Close()
            'SaveStream.Write(outp, 0, outp.Count * 2)
            'FilePut(f, outp, SoundWAVHeaderSize() + 1)
            FileOpen(f, FileName, OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
            SoundWriteWAVHeader(f, iFreq, 2)
            FileClose(f)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ex.TargetSite.Name)
            x = x
        End Try
    End Sub

    Public Sub SoundCombineSaveMono(ByVal xLeft() As Short, ByVal xRight() As Short, ByVal FileName As String, ByVal iFreq As Integer)
        If UBound(xLeft) > UBound(xRight) Then
            ReDim Preserve xRight(UBound(xLeft))
        ElseIf UBound(xLeft) < UBound(xRight) Then
            ReDim Preserve xLeft(UBound(xRight))
        End If
        Dim outp() As Short
        Dim x As Integer
        Dim y As Integer
        ReDim outp((UBound(xLeft) * 2) - 1)
        For x = 0 To UBound(outp) Step 2
            outp(x) = xLeft(y)
            outp(x + 1) = xRight(y)
            y += 1
        Next
        Dim f As Integer = FreeFile()
        FileOpen(f, FileName, OpenMode.Binary, OpenAccess.Write, OpenShare.Shared)
        Try
            FilePut(f, outp, SoundWAVHeaderSize() + 1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ex.TargetSite.Name)
            x = x
        End Try
        SoundWriteWAVHeader(f, iFreq, 2)
        FileClose(f)
    End Sub

    Public Sub SoundUpsample(ByVal xSource() As Byte, ByRef xDest() As Short, ByVal bFlipSign As Boolean, ByVal iBits As Integer, ByVal iChannels As Integer, ByVal iFreq As Integer, ByVal VolL As Double, ByVal VolR As Double)
        Dim fourbitmap() As Byte = {&H0, &H1, &H2, &H4, &H8, &H10, &H20, &H40, 0, &HC0, &HE0, &HF0, &HF8, &HFC, &HFE, &HFF}
        Dim LastSample As Integer
        Dim ThisSample As Integer
        Dim Outp() As Short
        Dim AdjustIndex As Byte
        Dim FinalSize As Integer
        Dim o As Integer
        Dim x As Integer
        Dim y As Integer
        If (iChannels > 2 Or iChannels < 1) Or (iBits <> 4 And iBits <> 8 And iBits <> 16) Then
            Exit Sub
        End If
        If iBits = 16 Then
            'can't have an odd byte count when using 16 bit
            If (UBound(xSource) + 1) Mod 2 <> 0 Then
                ReDim Preserve xSource(UBound(xSource) + 1)
            End If
        End If
        FinalSize = UBound(xSource) + 1
        If iChannels = 1 Then
            FinalSize *= 2
        End If
        If iBits = 4 Then
            FinalSize *= 2
        End If
        ReDim Outp(FinalSize - 1)
        If iBits = 4 Then
            LastSample = 0
            If iChannels = 1 Then
                For x = 0 To UBound(xSource)
                    AdjustIndex = xSource(x)
                    For y = 0 To 1
                        'convert 4 delta bits to 8 bits
                        ThisSample = (LastSample + fourbitmap(AdjustIndex And 15)) And 255
                        LastSample = ThisSample
                        'convert 8 bits to 16 bits
                        Outp(o) = ThisSample
                        Outp(o) <<= 8
                        Outp(o + 1) = Outp(o)
                        o += 2
                        AdjustIndex >>= 4
                    Next
                Next
            End If
        ElseIf iBits = 8 Then
            If iChannels = 1 Then
                For x = 0 To UBound(xSource)
                    Outp(o) = CShort(xSource(x)) << 8
                    Outp(o + 1) = Outp(o)
                    o += 2
                Next
            Else
                For x = 0 To UBound(xSource)
                    Outp(o) = CShort(xSource(x)) << 8
                Next
            End If
            'ReduceVolume = True
        ElseIf iBits = 16 Then
            If iChannels = 1 Then
                If bFlipSign Then
                    For x = 0 To UBound(xSource) Step 2
                        Outp(o) = CShort(xSource(x))
                        Outp(o) = Outp(o) Or (CShort(xSource(x + 1)) << 8)
                        If (xSource(x + 1) And 128) Then
                            Outp(o) = -(Outp(o) And &H7FFF)
                        End If
                        Outp(o + 1) = Outp(o)
                        o += 2
                    Next
                Else
                    For x = 0 To UBound(xSource) Step 2
                        Outp(o) = CShort(xSource(x))
                        Outp(o) = Outp(o) Or (CShort(xSource(x + 1)) << 8)
                        Outp(o + 1) = Outp(o)
                        o += 2
                    Next
                End If
            Else
                If bFlipSign Then
                    For x = 0 To UBound(xSource) Step 2
                        Outp(o) = CShort(xSource(x))
                        Outp(o) = Outp(o) Or (CShort(xSource(x + 1)) << 8)
                        If (xSource(x + 1) And 128) Then
                            Outp(o) = -(Outp(o) And &H7FFF)
                        End If
                        o += 1
                    Next
                Else
                    For x = 0 To UBound(xSource) Step 2
                        Outp(o) = CShort(xSource(x))
                        Outp(o) = Outp(o) Or (CShort(xSource(x + 1)) << 8)
                        o += 1
                    Next
                End If
            End If
        End If
        If VolL > 1 Or VolR > 1 Then
            For x = 0 To UBound(Outp) Step 2
                ThisSample = CInt(CDbl(Outp(x)) * VolL)
                If ThisSample > 32767 Then ThisSample = 32767
                If ThisSample < -32768 Then ThisSample = -32768
                Outp(x) = CShort(ThisSample)
                ThisSample = CInt(CDbl(Outp(x + 1)) * VolR)
                If ThisSample > 32767 Then ThisSample = 32767
                If ThisSample < -32768 Then ThisSample = -32768
                Outp(x + 1) = CShort(ThisSample)
            Next
        ElseIf VolL < 1 Or VolR < 1 Then
            For x = 0 To UBound(Outp) Step 2
                Outp(x) = CShort(CDbl(Outp(x)) * VolL)
                Outp(x + 1) = CShort(CDbl(Outp(x + 1)) * VolR)
            Next
        End If
        xDest = Outp
    End Sub

    Public Sub SoundDecodeVAGStream(ByVal sTargetFileName As String, ByVal SourceStream As IO.Stream, ByVal iChannels As Integer, ByVal iBlockSize As Integer, ByVal iOffset As Long, ByVal iFreq As Integer, ByVal iDecodeType As SoundDecodeType, ByVal iDecodeParam As Integer, Optional ByVal fVolumeL As Single = 1, Optional ByVal fVolumeR As Single = 1, Optional ByVal iMaxLength As Long = -1)
        Dim ConvertMonoToStereo As Boolean = (iChannels = 1)
        If iMaxLength = 0 Then
            Exit Sub
        End If
        If iChannels <= 0 Then
            Exit Sub
        End If
        SourceStream.Position = iOffset
        iOffset += 1

        Dim x As Integer
        Dim y As Integer
        Dim e As Boolean
        Dim l As Integer
        Dim o As Integer
        Dim z As Integer
        Dim dec1 As Byte = (iDecodeParam And 255)
        Dim dec2 As Byte = ((iDecodeParam >> 8) And 255)
        Dim dec3 As Byte = ((iDecodeParam >> 16) And 255)
        Dim dec4 As Byte = ((iDecodeParam >> 24) And 255)
        Dim bDecode As Boolean
        Dim pb As Boolean = (iMaxLength > 0)
        Dim filter As Byte
        Dim magnitude As Byte
        Dim InBuffer() As Byte
        Dim OutBuffer() As Short
        Dim MaxOffset As Long = iOffset
        Dim AutoFindLength As Boolean = (iMaxLength < 0)
        Dim ChannelBuffers() As IO.MemoryStream
        Dim DecodeBuffers() As SoundDecodeBuffer

        ReDim ChannelBuffers(0 To iChannels - 1)
        ReDim DecodeBuffers(UBound(ChannelBuffers))
        For x = 0 To iChannels - 1
            If iMaxLength > 0 Then
                ChannelBuffers(x) = New IO.MemoryStream(iMaxLength)
            Else
                ChannelBuffers(x) = New IO.MemoryStream
            End If
            DecodeBuffers(x) = New SoundDecodeBuffer
        Next
        If iMaxLength > 0 Then
            MaxOffset += iMaxLength
        End If

        Dim d As Integer
        Dim d1 As Integer
        Dim d2 As Integer
        Dim db As Byte
        Dim c As Integer
        Dim p1 As Integer
        Dim p2 As Integer
        'Seek(iFileNumber, iOffset)

        'read compressed audio
        ReDim InBuffer(0 To (iBlockSize - 1))
        Do While (Not e) And ((iOffset < MaxOffset) Or (AutoFindLength))
            For y = 0 To iChannels - 1
                'FileGet(iFileNumber, InBuffer)
                SourceStream.Read(InBuffer, 0, iBlockSize)
                iOffset += iBlockSize
                If iDecodeType = SoundDecodeType.BMDXgeneric Then
                    For x = 0 To iBlockSize - 1 Step 16
                        If InBuffer(x + 1) = 4 Then
                            dec2 = 0
                        End If
                        InBuffer(x) = InBuffer(x) Xor dec1
                        InBuffer(x + 1) = InBuffer(x + 1) Xor dec2
                        InBuffer(x + 2) = (CShort(InBuffer(x + 2)) - dec3) And 255
                        InBuffer(x + 3) = (CShort(InBuffer(x + 3)) - dec4) And 255
                    Next
                End If
                For x = 0 To iBlockSize - 1 Step 16
                    If (InBuffer(x + 1) And 1) Then
                        e = True
                    End If
                Next x
                ChannelBuffers(y).Write(InBuffer, 0, iBlockSize)
            Next
        Loop
        c = 0

        'convert compressed audio (only first two channels)
        If iChannels > 2 Then
            iChannels = 2
        End If
        ReDim InBuffer(0 To 15)
        For y = 0 To iChannels - 1
            ChannelBuffers(y).Position = 0
            l = ((ChannelBuffers(y).Length \ 16) * 28)
            ReDim DecodeBuffers(y).Data(l - 1)
            o = 0
            d = 0
            p2 = 0
            p1 = 0
            c = 0
            bDecode = True
            z = ChannelBuffers(y).Length
            Do While o < z
                ChannelBuffers(y).Read(InBuffer, 0, 16)
                o += 16
                filter = InBuffer(0) >> 4
                magnitude = InBuffer(0) And 15
                If (filter > 4) OrElse (magnitude > 12) Then
                    magnitude = 12
                    filter = 0
                End If
                If bDecode Then
                    For x = 2 To 15

                        db = InBuffer(x)
                        d1 = (CInt(db) And &HF) << (12 + 16)
                        d2 = (CInt(db) And &HF0) << (8 + 16)
                        d1 >>= magnitude + 16
                        d2 >>= magnitude + 16

                        c = d1 + (((p1 * VAGCoEff(filter, 0)) + (p2 * VAGCoEff(filter, 1))) >> 6)
                        If c < -32768 Then
                            c = -32768
                        ElseIf c > 32767 Then
                            c = 32767
                        End If
                        p2 = p1
                        p1 = c
                        DecodeBuffers(y).Data(d) = CShort(c)
                        d += 1

                        c = d2 + (((p1 * VAGCoEff(filter, 0)) + (p2 * VAGCoEff(filter, 1))) >> 6)
                        If c < -32768 Then
                            c = -32768
                        ElseIf c > 32767 Then
                            c = 32767
                        End If
                        p2 = p1
                        p1 = c
                        DecodeBuffers(y).Data(d) = CShort(c)
                        d += 1

                    Next
                    If (InBuffer(1) And 1) Then
                        bDecode = False
                    End If
                End If
            Loop
            ChannelBuffers(y) = Nothing
        Next

        'convert mono to stereo
        If ConvertMonoToStereo Then
            iChannels = 2
            ReDim Preserve DecodeBuffers(0 To 1)
            ReDim DecodeBuffers(1).Data(UBound(DecodeBuffers(0).Data))
            Array.ConstrainedCopy(DecodeBuffers(0).Data, 0, DecodeBuffers(1).Data, 0, UBound(DecodeBuffers(0).Data) + 1)
        End If

        'process velocity
        If (fVolumeL <> 1) OrElse (fVolumeR <> 1) Then
            For y = 0 To iChannels - 2
                If fVolumeL <> 1 Then
                    For x = 0 To UBound(DecodeBuffers(y).Data)
                        c = CInt(DecodeBuffers(y).Data(x))
                        c *= fVolumeL
                        If c < -32768 Then
                            c = -32768
                        ElseIf c > 32767 Then
                            c = 32767
                        End If
                        DecodeBuffers(y).Data(x) = CShort(c)
                    Next
                End If
                y += 1
                If fVolumeR <> 1 Then
                    For x = 0 To UBound(DecodeBuffers(y).Data)
                        c = CInt(DecodeBuffers(y).Data(x))
                        c *= fVolumeR
                        If c < -32768 Then
                            c = -32768
                        ElseIf c > 32767 Then
                            c = 32767
                        End If
                        DecodeBuffers(y).Data(x) = CShort(c)
                    Next
                End If
            Next
        End If

        'combine channels to stereo
        l = UBound(DecodeBuffers(0).Data) + 1
        y = 0
        ReDim OutBuffer(0 To (l * 2) - 1)
        For x = 0 To UBound(OutBuffer) Step 2
            OutBuffer(x) = DecodeBuffers(0).Data(y)
            OutBuffer(x + 1) = DecodeBuffers(1).Data(y)
            y += 1
        Next

        SoundSave(OutBuffer, sTargetFileName, iFreq)

    End Sub

    'conversion from the REALLY FUCKING ARBITRARY frequencies in the keysound list
    Public Function SoundConvertFrequency(ByVal inFreq As Integer) As Integer
        Dim x As Integer
        Dim sl As Double
        Dim ThisFreq As Integer
        Dim ThisVal As Integer
        Dim NextFreq As Integer
        Dim CalcFreq As Integer
        Dim NextVal As Integer
        'search for an exact match
        For x = 0 To UBound(FreqTable) Step 2
            If (inFreq = FreqTable(x + 1)) Then
                Return FreqTable(x)
                Exit Function
            End If
        Next
        'calculate an approximate frequency if we don't have one
        For x = 2 To UBound(FreqTable) Step 2
            If inFreq < FreqTable(x + 1) Or (x > UBound(FreqTable) - 2) Then
                ThisVal = FreqTable(x - 1)
                ThisFreq = FreqTable(x - 2)
                NextVal = FreqTable(x + 1)
                NextFreq = FreqTable(x)

                sl = (NextFreq - ThisFreq) / (NextVal - ThisVal)
                CalcFreq = ((inFreq - ThisVal) * sl) + ThisFreq
                Exit For
            End If
        Next
        If CalcFreq < 8000 Then
            CalcFreq = 8000
        End If
        DebugLog("SoundConvertFrequency", "An exact frequency match could not be found for this keysound.", "The guessed value is interpolated between known matches." & vbCrLf & "Input value: " & inFreq & vbCrLf & "Guessed frequency: " & CalcFreq)
        Return CalcFreq
    End Function

End Module
