Option Explicit On
Module modDebug
    Public Structure DebugInfoType
        Public sModule As String
        Public sError As String
        Public sInfo As String
        Public sTime As String
        Public sDate As String
    End Structure

    Public DebugInfo() As DebugInfoType
    Public DebugIndex As Integer = -1

    'report an error
    Public Sub DebugLog(ByVal sModule As String, ByVal sError As String, Optional ByVal sInfo As String = "")
        DebugIndex += 1
        ReDim Preserve DebugInfo(DebugIndex)
        Debug.Print(sModule & ": " & sError)
        With DebugInfo(DebugIndex)
            .sError = sError
            .sInfo = sInfo
            .sModule = sModule
            .sTime = (Today & " " & TimeOfDay)
        End With
    End Sub

End Module
