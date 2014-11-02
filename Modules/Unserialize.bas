Attribute VB_Name = "Unserialize"
Public Function unserialize(ByVal str As String) As Variant
    Dim retval() As Variant, Index As Integer, i As Integer, inpdata As String
    inpdata = str
    
    ' Checks used split character for Decimal (depends from country, dot or comma)
    Dim commaused As Boolean
    If CStr(CDec("5.44")) = "544" Then
        commaused = True
    Else
        commaused = False
    End If
    
    ' Used at String, Integer, Array
    Dim nr As String, b As String
    
    ' Used at Array
    Dim arraydata As String, arrayend As Boolean, arraydepth As Integer, initstate As Boolean
    
    If Len(inpdata) > 32767 Then
        ' Data may not be longer then 32767 characters. Otherwise it will raise an overflow error
        Error 6
    End If
    
    i = 0
    Do
        i = i + 1
        Select Case UCase(Mid(inpdata, i, 1))
            Case "A"    ' Array pointer (in this routine called Subarray)
                ' New value in array
                Index = Index + 1: ReDim Preserve retval(1 To Index)
                
                Do ' Loops until next { to get to the array start, I don't parse the value count, since it's not required data
                    i = i + 1
                    b = Mid(inpdata, i, 1)
                    If b <> "{" Then nr = nr & b
                Loop Until b = "{" Or i >= Len(inpdata)
                arraydata = "": b = "": arrayend = False: initstate = True
                
                Do ' Loops until array is ended
                    i = i + 1
                    b = Mid(inpdata, i, 1)
                    arraydata = arraydata & b
                    ' Checks for strings
                    If initstate Then
                        If UCase(b) = "S" Then ' String in subarray, can contain } so be careful
                            i = i + 1: arraydata = arraydata & ":" ' Expects a : right after the string-declaration
                            
                            ' Almost exact copy of string functionality
                            nr = "": b = ""
                            Do ' Loops until next : to get the string length
                                i = i + 1
                                b = Mid(inpdata, i, 1)
                                arraydata = arraydata & b
                                If b <> ":" Then nr = nr & b
                            Loop Until b = ":" Or i >= Len(inpdata)
                            arraydata = arraydata & Chr(34) & Mid(inpdata, i + 2, Val(nr)) & Chr(34) & ";"
                            i = i + Val(nr) + 3 ' 3 = two times " and one time ;
                            
                        End If
                        initstate = False
                    Else
                        If b = ";" Then initstate = True
                    End If
                    ' Subsub arrays will be handled here
                    If b = "{" Then arraydepth = arraydepth + 1
                    If b = "}" Then
                        If arraydepth = 0 Then arrayend = True Else arraydepth = arraydepth - 1
                    End If
                    If i >= Len(inpdata) Then arrayend = True  ' Exeption
                Loop Until arrayend
                
                arraydata = Left(arraydata, Len(arraydata) - 1) ' Removes last }
                retval(Index) = unserialize(arraydata)
            Case "S"    ' String pointer
                ' New value in array
                Index = Index + 1: ReDim Preserve retval(1 To Index)
                i = i + 1 ' Expects a : right after the string-declaration
                nr = "": b = ""
                Do ' Loops until next : to get the string length
                    i = i + 1
                    b = Mid(inpdata, i, 1)
                    If b <> ":" Then nr = nr & b
                Loop Until b = ":" Or i >= Len(inpdata)
                retval(Index) = Mid(inpdata, i + 2, Val(nr))
                i = i + Val(nr) + 3 ' 3 = two times " and one time ;
            Case "B"    ' Boolean pointer
                ' New value in array
                Index = Index + 1: ReDim Preserve retval(1 To Index)
                i = i + 1 ' Expects a : right after the boolean-declaration
                i = i + 1 ' Boolean-value itself
                If Mid(inpdata, i, 1) = "1" Then retval(Index) = True Else retval(Index) = False
                i = i + 1 ' Expects a ; right after the boolean-value
            Case "I"    ' Integer pointer
                ' New value in array
                Index = Index + 1: ReDim Preserve retval(1 To Index)
                i = i + 1 ' Expects a : right after the integer-declaration
                nr = "": b = ""
                Do ' Loops until next ; to get the string length
                    i = i + 1
                    b = Mid(inpdata, i, 1)
                    If b <> ";" Then nr = nr & b
                Loop Until b = ";" Or i >= Len(inpdata)
                retval(Index) = CInt(nr)
            Case "D"    ' Double pointer
                ' New value in array
                Index = Index + 1: ReDim Preserve retval(1 To Index)
                i = i + 1 ' Expects a : right after the double-declaration
                nr = "": b = ""
                Do ' Loops until next ; to get the string length
                    i = i + 1
                    b = Mid(inpdata, i, 1)
                    If b <> ";" Then nr = nr & b
                Loop Until b = ";" Or i >= Len(inpdata)
                
                If commaused Then nr = Replace(nr, ".", ",")
                retval(Index) = CDec(nr)
            Case "N"    ' Null pointer
                ' New value in array
                Index = Index + 1: ReDim Preserve retval(1 To Index)
                retval(Index) = Null
                i = i + 1 ' Expects a ; at the end of the Null
            Case Else
                ' The string contains unparsable values
                Error 93
        End Select
    Loop Until i >= Len(inpdata)
    
    unserialize = retval
End Function
