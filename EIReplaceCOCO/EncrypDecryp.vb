Public Class EncrypDecryp
    Public Function Decrypt(ByVal s As String, Optional ByVal key As String = "") As String
        Dim rStr As String = ""
        Dim i As Integer
        Dim ChkSum As Byte = 12
        Try
            If key = "" Then key = "oil"
            For i = 0 To s.Length - 1
                rStr &= Chr(Asc(s(i)) - 10)
            Next
            rStr = rStr.Remove(rStr.Length - (key.Length + 1))
            rStr = rStr.Remove(0, 1)
            Return rStr
        Catch ex As Exception
            Return s
        Finally
            rStr = Nothing
            i = Nothing
            ChkSum = Nothing
        End Try
    End Function

    Public Function base64Encode(ByVal pstr As String, Optional ByVal Password As String = "") As String
        Dim lencData_byte(pstr.Length) As Byte
        Dim lencData As String = ""

        Try
            lencData_byte = System.Text.Encoding.UTF8.GetBytes(pstr)
            lencData = Convert.ToBase64String(lencData_byte)

        Catch ex As Exception
            '
        End Try

        Return lencData

    End Function
End Class
