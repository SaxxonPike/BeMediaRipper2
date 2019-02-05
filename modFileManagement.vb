Option Explicit On
Module modFileManagement
    Public sDecoderInfo As String

    'loads a file into an array, returns TRUE if there was a problem
    Public Function FileLoadMemory(ByRef bTarget() As Byte, ByVal sFileName As String, Optional ByVal lOffset As Long = 0, Optional ByVal lLength As Long = -1) As Boolean
        FileLoadMemory = False
        'Dim f As Integer = FreeFile()
        If Not FileExists(sFileName) Then
            DebugLog("FileLoadMemory", "File does not exist.", sFileName)
            Return True
        End If
        If lLength = -1 Then 'if not specified, read to the end
            lLength = FileLen(sFileName) - lOffset
        End If
        Dim f As IO.Stream = IO.Stream.Null
        Try
            f = New IO.FileStream(sFileName, IO.FileMode.Open, IO.FileAccess.Read)
            ReDim bTarget(0 To lLength - 1)
            f.Position = lOffset
            f.Read(bTarget, 0, lLength)
        Catch ex As Exception
            DebugLog("FileLoadMemory", ex.Message, sFileName)
            FileLoadMemory = True
            f.Close()
            Exit Function
        End Try
        f.Close()
        FileLoadMemory = False
    End Function

    'opens a file for read access, returns file number (or 0 if there was a problem)
    Public Function FileLoad(ByVal sFileName As String, Optional ByVal lOffset As Long = 0) As IO.Stream 'Integer
        FileLoad = IO.Stream.Null
        If Not FileExists(sFileName) Then
            DebugLog("FileLoad", "File does not exist.", sFileName)
            Exit Function
        End If
        Try
            FileLoad = New IO.FileStream(sFileName, IO.FileMode.Open, IO.FileAccess.Read)
        Catch ex As Exception
            FileLoad = IO.Stream.Null
            DebugLog("FileLoad", ex.Message, sFileName)
            Exit Function
        End Try
    End Function

    'saves an array to a file
    Public Sub FileSaveMemory(ByRef bSource() As Byte, ByVal sFileName As String, Optional ByVal lOffset As Long = 0)
        Dim f As Integer = FreeFile()
        Try
            FileOpen(f, sFileName, OpenMode.Binary, OpenAccess.Write, OpenShare.LockWrite)
            FilePutObject(f, bSource, lOffset + 1)
        Catch ex As Exception
            DebugLog("FileSaveMemory", ex.Message, sFileName)
        End Try
        FileClose(f)
    End Sub

    'creates a folder
    Public Function FileCreateFolder(ByVal sFolderName As String) As Boolean
        If My.Computer.FileSystem.DirectoryExists(sFolderName) = True Then
            Return False
        End If
        Try
            FileSystem.MkDir(sFolderName)
            Return True
        Catch ex As Exception
            DebugLog("FileCreateFolder", ex.Message, sFolderName)
            Return False
        End Try
    End Function

    'opens a file for write access, returns file number (or 0 if there was a problem)
    Public Function FileSave(ByVal sFileName As String, Optional ByVal lOffset As Long = 0) As Integer
        FileSave = FreeFile()
        Try
            FileOpen(FileSave, sFileName, OpenMode.Binary, OpenAccess.Write, OpenShare.LockWrite, lOffset + 1)
        Catch ex As Exception
            FileSave = 0
            DebugLog("FileSave", ex.Message, sFileName)
        End Try
    End Function

    'check to see if a file exists. returns TRUE if it does
    Public Function FileExists(ByVal sFileName As String) As Boolean
        FileExists = (Dir(sFileName) <> "")
    End Function

    'returns a single line of text from a memory array
    Public Function FileGetLine(ByRef bSource() As Byte, ByRef iPtr As Integer) As String
        Dim x As Integer = UBound(bSource)
        FileGetLine = ""
        Do While iPtr < x
            If bSource(iPtr) = 13 Then
                iPtr += 1
                Exit Do
            ElseIf (bSource(iPtr) >= 32) OrElse (bSource(iPtr) = 9) OrElse (bSource(iPtr) = 0) Then
                FileGetLine &= Chr(bSource(iPtr))
            End If
            iPtr += 1
        Loop
    End Function

    'returns the application's path
    Public Function FileAppPath() As String
        FileAppPath = System.AppDomain.CurrentDomain.BaseDirectory
    End Function


End Module
