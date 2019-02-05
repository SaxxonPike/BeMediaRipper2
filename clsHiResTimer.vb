Imports System.Runtime.InteropServices
Imports System.Security

Public Class clsHiResTimer

    <DllImport("kernel32.dll"), SuppressUnmanagedCodeSecurity()> _
    Private Shared Function QueryPerformanceCounter(ByRef lpPerformanceCount As Long) As Boolean
    End Function

    <DllImport("kernel32.dll"), SuppressUnmanagedCodeSecurity()> _
    Private Shared Function QueryPerformanceFrequency(ByRef lpPerformanceFreq As Long) As Boolean
    End Function

    Private Shared freq As Long
    Shared Sub New()
        QueryPerformanceFrequency(freq)
    End Sub

    Private startTime, endTime As Long
    Private _duration As Double

    Public Sub Start()
        QueryPerformanceCounter(startTime)
    End Sub

    Public Function [Stop]() As Double
        QueryPerformanceCounter(endTime)
        _duration = ((endTime - startTime) / freq)
        Return _duration
    End Function

    Public ReadOnly Property Duration() As Double
        Get
            Return _duration
        End Get
    End Property

    Public Sub Update()
        QueryPerformanceCounter(endTime)
        _duration = ((endTime - startTime) / freq)
    End Sub
End Class
