Module modDragDrop
    Public Sub DragEnterFiles(ByRef sender As Object, ByRef e As System.Windows.Forms.DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Public Function DragDropFiles(ByRef sender As Object, ByRef e As System.Windows.Forms.DragEventArgs) As String()
        ReDim DragDropFiles(0)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            DragDropFiles = e.Data.GetData(DataFormats.FileDrop)
        End If
    End Function
    Public Function DragDropFiles(ByRef sender As Object, ByRef e As System.Windows.Forms.DragEventArgs, ByVal Index As Integer) As String
        Dim MyFiles(0)
        DragDropFiles = ""
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            If UBound(MyFiles) >= Index Then
                DragDropFiles = MyFiles(Index)
            End If
        End If
    End Function

End Module
