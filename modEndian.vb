Module modEndian
    Public Function Endian(ByVal d2 As Int16) As Int16
        Return (d2 >> 8) Or ((d2 And &HFF) << 8)
    End Function
    Public Function Endian(ByVal d4 As Int32) As Int32
        Endian = (d4 >> 24) And &HFF%
        Endian += (d4 >> 8) And &HFF00%
        Endian += (d4 << 8) And &HFF0000%
        Endian = Endian Or ((d4 << 24) And &HFF000000%)
    End Function
    Public Function Endian(ByVal d8 As Int64) As Int64
        Debug.Print(Hex(d8))
        Endian = (d8 >> 56) And &HFF&
        Endian += (d8 >> 40) And &HFF00&
        Endian += (d8 >> 24) And &HFF0000&
        Endian += (d8 >> 8) And &HFF000000&
        Endian += (d8 << 8) And &HFF00000000&
        Endian += (d8 << 24) And &HFF0000000000&
        Endian += (d8 << 40) And &HFF000000000000&
        Endian = Endian Or ((d8 << 56) And &HFF00000000000000&)
    End Function
End Module
