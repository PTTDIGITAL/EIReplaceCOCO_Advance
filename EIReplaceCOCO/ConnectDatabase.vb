Public Class ConnectDatabase

    'Dim Server As String = "10.195.2.177"
    'Dim Password As String = "pTT!CT01"
    'Public INIFile As String = Application.StartupPath & "\config.ini"
    Public ConnStr As String = getConnectionString()

    Public Function getConnectionString() As String

        Dim Server As String = "."
        'Dim Server As String = "10.195.2.205"

        Dim Database As String = "POSDB"
        'Dim Database As String = "POSDB_Export"

        Dim Username As String = "sa"

        Dim Password As String = "1qaz@WSX"
        'Dim Password As String = "pTT!CT01"


        'Dim ini As New IniReader(INIFile)
        'ini.Section = "Setting"
        Return "Data Source=" & Server & ";Initial Catalog=" & Database & ";User ID=" & Username & ";Password=" & Password & ";Connect Timeout=1;"
    End Function



End Class
