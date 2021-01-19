Module Module2
    Public P_DtView1 As DataView

    Public Function MidB(ByVal key1, ByVal key2, ByVal key3) As String
        Dim loop1, Len1, LenB1 As Integer
        Dim rtn As String

        For loop1 = 1 To key2 + key3 + 1
            Len1 = Len(Left(key1, loop1))
            LenB1 = System.Text.Encoding.GetEncoding(932).GetByteCount(Left(key1, loop1))

            If LenB1 >= key2 Then
                If LenB1 < key2 + key3 Then
                    rtn = rtn & Mid(key1, Len1, 1)
                    If System.Text.Encoding.GetEncoding(932).GetByteCount(rtn) = key3 Then
                        Return rtn
                    End If
                Else
                    Return rtn
                End If
            End If
        Next
    End Function

    Public Function IsDate(ByVal YearPart As Object, ByVal MonthPart As Object, ByVal DayPart As Object) As Boolean

        If IsNumeric(YearPart) And IsNumeric(MonthPart) And IsNumeric(DayPart) Then
            If CInt(YearPart) > 2100 OrElse _
              CInt(MonthPart) > DateTime.MaxValue.Month OrElse _
              CInt(DayPart) > DateTime.MaxValue.Day OrElse _
              CInt(YearPart) < 1900 OrElse _
              CInt(MonthPart) < DateTime.MinValue.Month OrElse _
              CInt(DayPart) < DateTime.MinValue.Day Then
                Return False
            Else
                Select Case CInt(MonthPart)
                    Case 1, 3, 5, 7, 8, 10, 12

                    Case 4, 6, 9, 11
                        If CInt(DayPart) > 30 Then
                            Return False
                        End If
                    Case 2
                        If DateTime.IsLeapYear(CInt(YearPart)) Then
                            If CInt(DayPart) > 29 Then
                                Return False
                            End If
                        Else
                            If CInt(DayPart) > 28 Then
                                Return False
                            End If
                        End If
                End Select
            End If
            Return True
        End If

    End Function

    Public Function numeric_check(ByVal i_data)
        Dim i As Integer

        For i = 1 To Len(i_data)
            Select Case Mid(i_data, i, 1)
                Case Is = "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                Case Else
                    Return "NG" : Exit Function
            End Select
        Next
        Return "OK"
    End Function

End Module
