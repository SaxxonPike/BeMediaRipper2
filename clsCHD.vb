Imports System.Runtime.InteropServices 'Used for dll importing
Imports System.IO.Compression

Public Class clsCHD
    'Much of this source was ported from the official MAME source code by SaxxonPike.
    'Portions (C) Nicola Salmoria and the MAME Team. (http://www.mamedev.org)

    'specific to this class ---------------------------------------------------------------------------

    Private Structure CHDcachetype
        Public dat() As Byte
        Public offs As Long
        Public hunk As Integer
    End Structure

    Private Shared CHDFileNumber As Integer = -1
    Private Shared CHDHeader As CHD_HEADER
    Private Shared CHDMap() As MAP_ENTRY
    Private Shared CHDDecBuffer() As Byte
    Private Shared CHDcache(0 To 511) As CHDcachetype
    Private Shared CHDcacheStack As Integer
    Public Shared CHDIdent As String

    Private Shared Function Read64M(Optional ByVal offs As Long = -1) As Int64
        Dim r As Int64
        If offs >= 0 Then
            FileGet(CHDFileNumber, r, offs + 1)
        Else
            FileGet(CHDFileNumber, r)
        End If
        Return Endian(r)
    End Function
    Private Shared Function Read64(Optional ByVal offs As Long = -1) As Int64
        Read64 = 0
        If offs >= 0 Then
            FileGet(CHDFileNumber, Read64, offs + 1)
        Else
            FileGet(CHDFileNumber, Read64)
        End If
    End Function
    Private Shared Function Read32M(Optional ByVal offs As Long = -1) As Int32
        Dim r As Int32
        If offs >= 0 Then
            FileGet(CHDFileNumber, r, offs + 1)
        Else
            FileGet(CHDFileNumber, r)
        End If
        Return Endian(r)
    End Function
    Private Shared Function Read32(Optional ByVal offs As Long = -1) As Int32
        Read32 = 0
        If offs >= 0 Then
            FileGet(CHDFileNumber, Read32, offs + 1)
        Else
            FileGet(CHDFileNumber, Read32)
        End If
    End Function
    Private Shared Function Read16M(Optional ByVal offs As Long = -1) As Int16
        Dim r As Int16
        If offs >= 0 Then
            FileGet(CHDFileNumber, r, offs + 1)
        Else
            FileGet(CHDFileNumber, r)
        End If
        Return Endian(r)
    End Function
    Private Shared Function Read16(Optional ByVal offs As Long = -1) As Int16
        Read16 = 0
        If offs >= 0 Then
            FileGet(CHDFileNumber, Read16, offs + 1)
        Else
            FileGet(CHDFileNumber, Read16)
        End If
    End Function
    Private Shared Function Read8(Optional ByVal offs As Long = -1) As Byte
        Read8 = 0
        If offs >= 0 Then
            FileGet(CHDFileNumber, Read8, offs + 1)
        Else
            FileGet(CHDFileNumber, Read8)
        End If
    End Function
    Private Shared Sub SelectCHD(ByVal sFileName As String)
        CloseCHD()
        CHDFileNumber = FreeFile()
        FileOpen(CHDFileNumber, sFileName, OpenMode.Binary, OpenAccess.Read, OpenShare.Shared)
    End Sub
    Private Shared Sub CloseCHD()
        If CHDFileNumber > -1 Then
            FileClose(CHDFileNumber)
            CHDFileNumber = -1
        End If
    End Sub

    'from CHD.H ---------------------------------------------------------------------------------------

    'header information
    Private Const CHD_HEADER_VERSION = 4
    Private Const CHD_V1_HEADER_SIZE = 76
    Private Const CHD_V2_HEADER_SIZE = 80
    Private Const CHD_V3_HEADER_SIZE = 120
    Private Const CHD_V4_HEADER_SIZE = 108
    Private Const CHD_MAX_HEADER_SIZE = CHD_V4_HEADER_SIZE

    'checksumming information
    Private Const CHD_MD5_BYTES = 16
    Private Const CHD_SHA1_BYTES = 20

    'CHD global flags
    Private Const CHDFLAGS_HAS_PARENT = 1
    Private Const CHDFLAGS_IS_WRITEABLE = 2
    Private Const CHDFLAGS_UNDEFINED = -4

    'compression types
    Private Const CHDCOMPRESSION_NONE = 0
    Private Const CHDCOMPRESSION_ZLIB = 1
    Private Const CHDCOMPRESSION_ZLIB_PLUS = 2
    Private Const CHDCOMPRESSION_AV = 3

    'error types
    Public Enum CHD_ERROR
        CHDERR_NONE
        CHDERR_NO_INTERFACE
        CHDERR_OUT_OF_MEMORY
        CHDERR_INVALID_FILE
        CHDERR_INVALID_PARAMETER
        CHDERR_INVALID_DATA
        CHDERR_FILE_NOT_FOUND
        CHDERR_REQUIRES_PARENT
        CHDERR_FILE_NOT_WRITEABLE
        CHDERR_READ_ERROR
        CHDERR_WRITE_ERROR
        CHDERR_CODEC_ERROR
        CHDERR_INVALID_PARENT
        CHDERR_HUNK_OUT_OF_RANGE
        CHDERR_DECOMPRESSION_ERROR
        CHDERR_COMPRESSION_ERROR
        CHDERR_CANT_CREATE_FILE
        CHDERR_CANT_VERIFY
        CHDERR_NOT_SUPPORTED
        CHDERR_METADATA_NOT_FOUND
        CHDERR_INVALID_METADATA_SIZE
        CHDERR_UNSUPPORTED_VERSION
        CHDERR_VERIFY_INCOMPLETE
        CHDERR_INVALID_METADATA
        CHDERR_INVALID_STATE
        CHDERR_OPERATION_PENDING
        CHDERR_NO_ASYNC_OPERATION
        CHDERR_UNSUPPORTED_FORMAT
    End Enum

    'extract header structure (NOT the on-disk header structure)
    Public Structure CHD_HEADER
        Public length As Int32
        Public version As Int32
        Public flags As Int32
        Public compression As Int32
        Public hunkbytes As Int32
        Public seclen As Int32
        Public totalhunks As Int32
        Public logicalbytes As Int64
        Public metaoffset As Int64
        Public md5() As Byte
        Public parentmd5() As Byte
        Public sha1() As Byte
        Public rawsha1() As Byte
        Public parentsha1() As Byte
        Public obsolete_cylinders As Int32
        Public obsolete_sectors As Int32
        Public obsolete_heads As Int32
        Public obsolete_hunksize As Int32
        Public Sub Init()
            ReDim md5(CHD_MD5_BYTES - 1)
            ReDim parentmd5(CHD_MD5_BYTES - 1)
            ReDim sha1(CHD_SHA1_BYTES - 1)
            ReDim rawsha1(CHD_SHA1_BYTES - 1)
            ReDim parentsha1(CHD_SHA1_BYTES - 1)
        End Sub
        Public Sub New(ByVal offs As Long)
            Init()
            length = Read32M(offs)
            version = Read32M()
            flags = Read32M()
            compression = Read32M()
            Select Case version
                Case 1
                    hunkbytes = Read32M()
                    totalhunks = Read32M()
                    obsolete_cylinders = Read32M()
                    obsolete_heads = Read32M()
                    obsolete_sectors = Read32M()
                    FileGet(CHDFileNumber, md5)
                    FileGet(CHDFileNumber, parentmd5)
                Case 2
                    hunkbytes = Read32M()
                    totalhunks = Read32M()
                    obsolete_cylinders = Read32M()
                    obsolete_heads = Read32M()
                    obsolete_sectors = Read32M()
                    FileGet(CHDFileNumber, md5)
                    FileGet(CHDFileNumber, parentmd5)
                    seclen = Read32M()
                Case 3
                    totalhunks = Read32M()
                    logicalbytes = Read64M()
                    metaoffset = Read64M()
                    FileGet(CHDFileNumber, md5)
                    FileGet(CHDFileNumber, parentmd5)
                    hunkbytes = Read32M()
                    FileGet(CHDFileNumber, sha1)
                    FileGet(CHDFileNumber, parentsha1)
                Case 4
                    totalhunks = Read32M()
                    logicalbytes = Read64M()
                    metaoffset = Read64M()
                    hunkbytes = Read32M()
                    FileGet(CHDFileNumber, sha1)
                    FileGet(CHDFileNumber, parentsha1)
                    FileGet(CHDFileNumber, rawsha1)
            End Select
        End Sub
    End Structure

    'from CHD.C ---------------------------------------------------------------------------------------

    Private Const MAP_STACK_ENTRIES = 512
    Private Const MAP_ENTRY_SIZE = 16
    Private Const OLD_MAP_ENTRY_SIZE = 8
    Private Const METADATA_HEADER_SIZE = 16
    Private Const CRCMAP_HASH_SIZE = 4095
    Private Const MAP_ENTRY_FLAG_TYPE_MASK = &HF
    Private Const MAP_ENTRY_FLAG_NO_CRC = &H10

    Private Const MAP_ENTRY_TYPE_INVALID = 0        'invalid type
    Private Const MAP_ENTRY_TYPE_COMPRESSED = 1     'compressed
    Private Const MAP_ENTRY_TYPE_UNCOMPRESSED = 2   'uncompressed
    Private Const MAP_ENTRY_TYPE_MINI = 3           'use offset as raw data
    Private Const MAP_ENTRY_TYPE_SELF_HUNK = 4      'same as another hunk in this file
    Private Const MAP_ENTRY_TYPE_PARENT_HUNK = 5    'same as a hunk in the parent file

    Private Const CHD_V1_SECTOR_SIZE = 512

    Private Const COOKIE_VALUE = &HBAADF00D&
    Private Const MAX_ZLIB_ALLOCS = 64
    Private Const END_OF_LIST_COOKIE = "EndOfListCookie"
    Private Const NO_MATCH = 0

    'a single map entry
    Public Structure MAP_ENTRY
        Public offset As Int64
        Public crc As Int32
        Public length As Int32
        Public flags As Byte
    End Structure

    'a single metadata entry
    Public Structure METADATA_ENTRY
        Public offset As Int64
        Public nextentry As Int64
        Public preventry As Int64
        Public length As Int32
        Public metatag As Int32
        Public flags As Byte
    End Structure

    Private Sub map_extract(ByVal offs As Long, ByRef mapentry As MAP_ENTRY)
        With mapentry
            .offset = Read64M(offs)
            .crc = Read32M()
            .length = Read16M() Or Read8()
            .flags = Read8()
        End With
    End Sub
    Private Sub map_extract_old(ByVal offs As Long, ByRef mapentry As MAP_ENTRY, ByVal hunkbytes As Int32)
        With mapentry
            .offset = Read64M(offs)
            .crc = 0
            .length = .offset >> 44
            .offset = (.offset And &HFFFFFFFFFFF&)
            .flags = MAP_ENTRY_FLAG_NO_CRC Or IIf(.length = hunkbytes, MAP_ENTRY_TYPE_UNCOMPRESSED, MAP_ENTRY_TYPE_COMPRESSED)
        End With
    End Sub
    Public Function chd_open_file(ByVal sFileName As String) As CHD_ERROR
        Dim ident As String = Space(8)
        Dim offs As Int64
        Dim mapoffs As Int64
        Dim x As Integer
        Dim mapbytes() As Byte = {}
        chd_open_file = CHD_ERROR.CHDERR_NONE
        offs = 8
        If Not FileExists(sFileName) Then
            Return CHD_ERROR.CHDERR_INVALID_PARAMETER
            Exit Function
        End If
        If (CHDHeader.flags And CHDFLAGS_HAS_PARENT) Then
            Return CHD_ERROR.CHDERR_REQUIRES_PARENT
            Exit Function
        End If
        SelectCHD(sFileName)
        FileGet(CHDFileNumber, ident, 1, True)
        If ident <> "MComprHD" Then
            CloseCHD()
            Return CHD_ERROR.CHDERR_INVALID_FILE
            Exit Function
        End If
        CHDHeader = New CHD_HEADER(offs)
        Dim entrysize As Integer = IIf(CHDHeader.version < 3, OLD_MAP_ENTRY_SIZE, MAP_ENTRY_SIZE)
        offs = CHDHeader.length
        mapoffs = 0
        If CHDHeader.totalhunks > 0 Then
            ReDim CHDMap(CHDHeader.totalhunks - 1)
            ReDim mapbytes((CHDHeader.totalhunks * entrysize) - 1)
            FileGet(CHDFileNumber, mapbytes, offs + 1, False)
            If entrysize = MAP_ENTRY_SIZE Then
                For x = 0 To CHDHeader.totalhunks - 1
                    With CHDMap(x)
                        .offset = DataMakeInt64(mapbytes(mapoffs + 7), mapbytes(mapoffs + 6), mapbytes(mapoffs + 5), mapbytes(mapoffs + 4), mapbytes(mapoffs + 3), mapbytes(mapoffs + 2), mapbytes(mapoffs + 1), mapbytes(mapoffs)) 'Read64M(offs)
                        .crc = DataMakeInt32(mapbytes(mapoffs + 11), mapbytes(mapoffs + 10), mapbytes(mapoffs + 9), mapbytes(mapoffs + 8))
                        .length = DataMakeInt32(mapbytes(mapoffs + 13), mapbytes(mapoffs + 12), mapbytes(mapoffs + 14), 0)
                        .flags = mapbytes(mapoffs + 15)
                    End With
                    mapoffs += entrysize
                Next
            Else
                x = x
            End If
            ReadHunkBytes(0, 32, mapbytes)
            DataByteSwap(mapbytes)
            CHDIdent = ""
            For x = 0 To UBound(mapbytes)
                CHDIdent &= Chr(mapbytes(x))
            Next
        End If
    End Function

    Public Function DecompressHunk(ByVal hunknum As Integer, ByRef Dest() As Byte) As Integer
        Dim decbuf() As Byte
        Dim x As Integer
        If hunknum < 0 Or hunknum > CHDHeader.totalhunks - 1 Then
            Return MAP_ENTRY_TYPE_INVALID
            Exit Function
        End If
        ReDim Dest(CHDHeader.hunkbytes - 1)
        With CHDMap(hunknum)
            DecompressHunk = (.flags And MAP_ENTRY_FLAG_TYPE_MASK)
            Select Case DecompressHunk
                Case MAP_ENTRY_TYPE_COMPRESSED
                    ReDim decbuf(.length - 1)
                    FileGet(CHDFileNumber, decbuf, .offset + 1, False)
                    Dim inpstream As New IO.MemoryStream(decbuf)
                    Dim decstream As New IO.Compression.DeflateStream(inpstream, CompressionMode.Decompress)
                    decstream.Read(Dest, 0, CHDHeader.hunkbytes)
                Case MAP_ENTRY_TYPE_UNCOMPRESSED
                    FileGet(CHDFileNumber, Dest, .offset + 1, False)
                Case MAP_ENTRY_TYPE_MINI
                    ReDim Dest(CHDHeader.hunkbytes - 1)
                    For x = 0 To CHDHeader.hunkbytes - 1 Step 8
                        Dest(x) = .offset And &HFF
                        Dest(x + 1) = (.offset >> 8) And &HFF
                        Dest(x + 2) = (.offset >> 16) And &HFF
                        Dest(x + 3) = (.offset >> 24) And &HFF
                        Dest(x + 4) = (.offset >> 32) And &HFF
                        Dest(x + 5) = (.offset >> 40) And &HFF
                        Dest(x + 6) = (.offset >> 48) And &HFF
                        Dest(x + 7) = (.offset >> 56) And &HFF
                    Next
                Case MAP_ENTRY_TYPE_SELF_HUNK
                    If .offset <> hunknum Then
                        DecompressHunk = DecompressHunk(.offset, Dest)
                    End If
                Case MAP_ENTRY_TYPE_PARENT_HUNK
                    'this code doesn't support parent hunks yet
                    'but no bemani game I know of even uses them anyway
                Case MAP_ENTRY_TYPE_INVALID
                Case Else
                    x = x
            End Select
        End With
    End Function

    Public Sub ExtractAllValidHunks(ByVal sFileName As String)
        Dim f As System.IO.FileStream = New IO.FileStream(sFileName, IO.FileMode.Create)
        Dim x As Integer
        Dim d() As Byte = {}
        For x = 0 To CHDHeader.totalhunks - 1
            Select Case (CHDMap(x).flags And MAP_ENTRY_FLAG_TYPE_MASK)
                Case MAP_ENTRY_TYPE_UNCOMPRESSED, MAP_ENTRY_TYPE_COMPRESSED, MAP_ENTRY_TYPE_SELF_HUNK
                    DecompressHunk(x, d)
                    f.Write(d, 0, d.Length)
            End Select
        Next
        f.Close()
    End Sub

    Public Sub ExtractAllHunks(ByVal sFileName As String)
        Dim f As System.IO.FileStream = New IO.FileStream(sFileName, IO.FileMode.Create)
        Dim x As Integer
        Dim d() As Byte = {}
        For x = 0 To CHDHeader.totalhunks - 1
            DecompressHunk(x, d)
            f.Write(d, 0, d.Length)
        Next
        f.Close()
    End Sub

    Public Function DecompressHunkByOffset(ByVal iOffset As Long, ByRef dat() As Byte)
        iOffset \= CHDHeader.hunkbytes
        Return DecompressHunk(CInt(iOffset), dat)
    End Function

    Public Function GetHunkFromOffset(ByVal iOffset As Long) As Integer
        Return iOffset \ CHDHeader.hunkbytes
    End Function

    Public Function HunkType(ByVal x As Integer) As Integer
        Return CHDMap(x).flags And MAP_ENTRY_FLAG_TYPE_MASK
    End Function

    Public Function HunkCount() As Integer
        Return UBound(CHDMap)
    End Function

    Public Function HunkSize() As Integer
        Return CHDHeader.hunkbytes
    End Function

    Public Sub ReadHunkBytes(ByVal iOffset As Long, ByVal iCount As Integer, ByRef dBytes() As Byte)
        Dim ThisHunk As Integer = iOffset \ CHDHeader.hunkbytes
        Dim NextHunkOffset As Long = ((iOffset \ CHDHeader.hunkbytes) + 1) * CHDHeader.hunkbytes
        Dim ThisCache As Integer = -1
        Dim OffsetWithinHunk As Integer = iOffset Mod CHDHeader.hunkbytes
        Dim x As Integer
        Dim y As Integer
        ReDim dBytes(iCount - 1)
        For x = 0 To UBound(CHDcache)
            If CHDcache(x).hunk = ThisHunk Then
                ThisCache = x
                Exit For
            End If
        Next
        If ThisCache = -1 Then
            ThisCache = LoadHunkToCache(ThisHunk)
        End If
        For x = 0 To iCount - 1
            If iOffset >= NextHunkOffset Then
                ThisHunk = iOffset \ CHDHeader.hunkbytes
                NextHunkOffset = ((iOffset \ CHDHeader.hunkbytes) + 1) * CHDHeader.hunkbytes
                OffsetWithinHunk = 0
                ThisCache = -1
                For y = 0 To UBound(CHDcache)
                    If CHDcache(y).hunk = ThisHunk Then
                        ThisCache = y
                        Exit For
                    End If
                Next
                If ThisCache = -1 Then
                    ThisCache = LoadHunkToCache(ThisHunk)
                End If
            End If
            dBytes(x) = CHDcache(ThisCache).dat(OffsetWithinHunk)
            OffsetWithinHunk += 1
            iOffset += 1
        Next
    End Sub

    Private Function LoadHunkToCache(ByVal hunk As Integer) As Integer
        DecompressHunk(hunk, CHDcache(CHDcacheStack).dat)
        LoadHunkToCache = CHDcacheStack
        CHDcacheStack += 1
        If CHDcacheStack > UBound(CHDcache) Then
            CHDcacheStack = 0
        End If
    End Function

    Public Sub New()
        Dim x As Integer
        For x = 0 To UBound(CHDcache)
            With CHDcache(x)
                .hunk = -1
                .offs = 0
                ReDim .dat(0 To 0)
            End With
        Next
    End Sub
End Class
