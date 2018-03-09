Imports Hnx8.ReadJEnc
Imports System.IO
Imports System.Text


Public Class EncodeResolver

    ' how to get encoding
    ' https://github.com/hnx8/ReadJEnc/blob/master/ReadJEnc_Readme.txt


    Public Shared Function GetEncoding(targetFile As String) As System.Text.Encoding

        Dim fi = New FileInfo(targetFile)
        Using reader = New FileReader(fi)

            Dim c = reader.Read(fi)
            Dim encode = c.GetEncoding()
            Return encode

        End Using

    End Function

End Class
