Imports System.Data.SqlClient
Imports System.IO

Public Class BackUpRestoreDB


    Dim ClsConn As New ConnectDatabase
    Dim ConnStr As String = ClsConn.ConnStr

    Public Function BackupDB(p As ProgressBar) As String
        Dim ret As String = ""
        Dim arr() As String = {"ENABLERDB", "POS_VIS", "FULLTAXIES", "POSDB"}
        Try
            Dim path As String = Application.StartupPath & "\DB"
            If Directory.Exists(path) Then
                For Each _file As String In Directory.GetFiles(path)
                    File.Delete(_file)
                Next
            Else
                Directory.CreateDirectory(path)
            End If

            For i As Integer = 0 To arr.Length - 1
                Dim s As String
                s = path & "\" & arr(i) & ".bak"
                Dim sql As String = "backup database " & arr(i) & " to disk='" & s & "'"

                Dim conn As New SqlConnection(ConnStr)
                conn.Open()
                Dim cmd As New SqlCommand
                With cmd
                    .CommandText = sql
                    .CommandType = CommandType.Text
                    .CommandTimeout = 500
                    .Connection = conn
                    .ExecuteNonQuery()
                End With
                p.Value = p.Value + 5
            Next
        Catch ex As Exception
            ret = "พบปัญหาในการ Backup Database  " & ex.ToString
        End Try
        Return ret
    End Function


    Public Function RestoreDB(p As ProgressBar) As String
        Dim ret As String = ""
        Try
            Dim path As String = Application.StartupPath & "\DB"
            If Directory.Exists(path) Then
                For Each _file As String In Directory.GetFiles(path)
                    Dim file_name As String = System.IO.Path.GetFileName(_file)
                    Dim str() As String = file_name.Split(".")
                    If str.Length > 0 Then
                        file_name = str(0)
                    End If

                    Dim dbpath As String = "C:\DB_DATA\"
                    If Not Directory.Exists(dbpath) Then
                        Directory.CreateDirectory(dbpath)
                    End If

                    If file_name.ToUpper = "POSDB" Then
                        file_name = "POSDB_COCO"
                    End If

                    Dim conn As New SqlConnection(ConnStr)
                    conn.Open()
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = conn
                    End With

                    Dim sql As String = ""
                    sql = "SELECT 'y' FROM sys.databases WHERE name = N'" & file_name.ToUpper & "'"
                    Dim da As New SqlDataAdapter(sql, ConnStr)
                    Dim dt As New DataTable
                    da.Fill(dt)

                    Dim data As String = ""
                    Dim log As String = ""
                    Select Case file_name.ToUpper
                        Case "ENABLERDB"
                            data = "ENBDATA"
                            log = "ENBLOG"
                        Case "POS_VIS"
                            data = "POS_VIS"
                            log = "POS_VIS_log"
                        Case "FULLTAXIES"
                            data = "FullTaxIES"
                            log = "FullTaxIES_log"
                        Case "POSDB_COCO"
                            data = "POSDB"
                            log = "POSDB_Log"
                    End Select


                    If dt.Rows.Count = 0 Then
                        sql = "RESTORE DATABASE " & file_name.ToUpper & "  FROM DISK = '" & _file & "'" & vbCrLf
                        sql &= " With Replace, RECOVERY, " & vbCrLf
                        sql &= " Move '" & data.ToUpper & "' TO 'C:\DB_DATA\" & data.ToUpper & ".mdf', " & vbCrLf
                        sql &= " Move '" & log.ToUpper & "' TO 'C:\DB_DATA\" & log.ToUpper & ".ldf'; " & vbCrLf

                    Else
                        sql = "use master" & vbCrLf
                        sql &= " alter database " & file_name & " " & vbCrLf
                        sql &= " Set offline With rollback immediate " & vbCrLf
                        sql &= " alter database " & file_name & " " & vbCrLf
                        sql &= " Set online With rollback immediate " & vbCrLf
                        cmd.CommandText = sql
                        cmd.CommandTimeout = 500
                        cmd.ExecuteNonQuery()

                        sql = " RESTORE DATABASE " & file_name.ToUpper & " FROM disk='" & _file & "' WITH REPLACE"

                    End If

                    cmd.CommandText = sql
                    cmd.CommandTimeout = 500
                    cmd.ExecuteNonQuery()

                    p.Value = p.Value + 5
                Next
            End If


        Catch ex As Exception
            ret = "พบปัญหาในการ Restore Database  " & ex.ToString
        End Try
        Return ret
    End Function


End Class
