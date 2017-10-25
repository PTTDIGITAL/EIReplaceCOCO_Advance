Public Class ConnectDatabase

    Public ConnStr As String = getConnectionString()
    Public ConnStrFullTax As String = getConnFullTaxIES()
    Public ConnStrPOSVIS As String = getConnPOSVIS()

    ReadOnly Property gPassword As String
        Get
            'Return "1qaz@WSX"
            Return "pTT!CT01"
        End Get
    End Property


    Public Function getConnectionString() As String

        Dim Server As String = "."
        Dim Database As String = "POSDB"
        'Dim Database As String = "POSDB_Export"

        Dim Username As String = "sa"
        Dim Password As String = gPassword
        Return "Data Source=" & Server & ";Initial Catalog=" & Database & ";User ID=" & Username & ";Password=" & Password & ";Connect Timeout=1;"
    End Function

    Public Function getConnFullTaxIES() As String
        Dim Server As String = "."
        Dim Database As String = "FullTaxIES"
        Dim Username As String = "sa"
        Dim Password As String = gPassword
        Return "Data Source=" & Server & ";Initial Catalog=" & Database & ";User ID=" & Username & ";Password=" & Password & ";Connect Timeout=1;"
    End Function

    Public Function getConnPOSVIS() As String
        Dim Server As String = "."
        Dim Database As String = "POS_VIS"
        Dim Username As String = "sa"
        Dim Password As String = gPassword
        Return "Data Source=" & Server & ";Initial Catalog=" & Database & ";User ID=" & Username & ";Password=" & Password & ";Connect Timeout=1;"
    End Function


End Class
