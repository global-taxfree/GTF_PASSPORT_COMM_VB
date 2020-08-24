Attribute VB_Name = "Module1"
Public Function BytesToString(bytes() As Byte) As String

    On Local Error GoTo ErrOut:

    

    Dim bytes2() As String, i As Long

    

    ReDim bytes2(UBound(bytes)) As String

    For i = 0 To UBound(bytes)

        bytes2(i) = Right("0" & Hex(bytes(i)), 2)

    Next

    BytesToString = Join(bytes2, "")

ErrOut:

End Function

