Attribute VB_Name = "Serialize"
Public Function serialize(arrdata As Variant) As String
    Dim retval As String, i As Integer
    
    Select Case TypeName(arrdata)
        Case "Decimal"
            retval = retval & "d:" & Replace(CStr(arrdata), ",", ".") & ";"
        Case "Integer"
            retval = retval & "i:" & CStr(arrdata) & ";"
        Case "Null"
            retval = retval & "N;"
        Case "Boolean"
            If arrdata Then
                retval = retval & "b:1;"
            Else
                retval = retval & "b:0;"
            End If
        Case "Variant()"
            retval = retval & "a:" & CStr(UBound(arrdata) - LBound(arrdata) + 1) & ":{"
            For i = LBound(arrdata) To UBound(arrdata)
                retval = retval & serialize(arrdata(i))
            Next i
            retval = retval & "}"
        Case Else ' String, but also all other types
            retval = retval & "s:" & CInt(Len(CStr(arrdata))) & ":" & Chr(34) & CStr(arrdata) & Chr(34) & ";"
    End Select
    
    serialize = retval
End Function
