Imports System.Data.SqlClient

Public Class GlobalFunction
    Dim ClsConn As New ConnectDatabase
    Dim ConnStr As String = ClsConn.ConnStr

    Public Function GetDateTime() As String
        Return DateTime.Now.ToString("yyyyMMdd H:mm:ss")
    End Function

    Public Function Update_APP_Config(CONFIG_KEY As String, CONFIG_VALUE As String) As Integer
        Dim sql As String = "SELECT * FROM APP_Config Where Config_Key = '" & CONFIG_KEY & "'"
        Dim da As New SqlDataAdapter(sql, ConnStr)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As Integer = 0
        If dt.Rows.Count > 0 Then
            sql = "Update APP_Config set CONFIG_VALUE = '" & CONFIG_VALUE & "' where Config_Key = '" & CONFIG_KEY & "' "
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.Text
                .Connection = conn
            End With
            ret = cmd.ExecuteNonQuery()
            conn.Close()
        End If
        Return ret
    End Function

    Public Function GET_PUMP_ID(HOSE_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT EH.Pump_ID FROM ENABLERDB.dbo.HOSES AS EH LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON EH.Grade_ID = PM.MAT_ID2 WHERE EH.HOSE_ID = " & HOSE_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("Pump_ID").ToString
        End If
        Return ret
    End Function

    Public Function GET_MAT_ID(HOSE_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT EH.Pump_ID, EH.Tank_ID, PM.MAT_ID FROM ENABLERDB.dbo.HOSES AS EH LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON EH.Grade_ID = PM.MAT_ID2 WHERE EH.HOSE_ID = " & HOSE_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("MAT_ID").ToString
        End If
        Return ret
    End Function

    Public Function GET_TANK_ID(HOSE_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT EH.Pump_ID, EH.Tank_ID, PM.MAT_ID FROM ENABLERDB.dbo.HOSES AS EH LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON EH.Grade_ID = PM.MAT_ID2 WHERE EH.HOSE_ID = " & HOSE_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("Tank_ID").ToString
        End If
        Return ret
    End Function

    Public Function GET_MAT_ID_TANK(TANK_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT PM.MAT_ID FROM ENABLERDB.dbo.Tanks AS ET LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON ET.Grade_ID = PM.MAT_ID2 WHERE ET.TANK_ID = " & TANK_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("MAT_ID").ToString
        End If
        Return ret
    End Function

    Public Function GET_TANK_NAME(TANK_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT ET.Tank_Name FROM ENABLERDB.dbo.Tanks AS ET LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON ET.Grade_ID = PM.MAT_ID2 WHERE ET.TANK_ID =" & TANK_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("Tank_Name").ToString
        End If
        Return ret
    End Function

    Public Function GET_TANK_NUMBER(TANK_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT ET.Tank_Number FROM ENABLERDB.dbo.Tanks AS ET LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON ET.Grade_ID = PM.MAT_ID2 WHERE ET.TANK_ID = " & TANK_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("Tank_Number").ToString
        End If
        Return ret
    End Function

    Public Function CallSPImportProduct() As String
        Try
            Dim sql As String = "sp_Import_Product_To_Inventory"
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.StoredProcedure
                .Connection = conn
                .ExecuteNonQuery()
            End With
            conn.Close()
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function

    Public Function UpdateDefaultProduct() As Integer
        Try
            Dim ret As Integer
            Dim sql As String = "Update PRODUCTS set ISSHOWINPOS = 0 , ISRECOMMEND=0 "
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.Text
                .Connection = conn
            End With
            ret = cmd.ExecuteNonQuery()
            conn.Close()
            Return ret
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function UpdateISSHOWINPOS() As Integer
        Try
            Dim ret As Integer
            Dim sql As String = "Update PRODUCTS set ISSHOWINPOS = 1 where ProductCode in (select MAT_ID from TBMATTERIAL_SITE)"
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.Text
                .Connection = conn
            End With
            ret = cmd.ExecuteNonQuery()
            conn.Close()
            Return ret
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function UpdateISRECOMMEND() As Integer
        Try
            Dim ret As Integer
            Dim sql As String = "Update PRODUCTS set ISRECOMMEND = 1 where ProductCode in (select MAT_ID from TBMAT_RECOMMEND)"
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.Text
                .Connection = conn

            End With
            ret = cmd.ExecuteNonQuery()
            conn.Close()
            Return ret
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function CallSPInitialLUBE() As String
        Try
            Dim sql As String = "sp_Initial_LUBE_Stock_Inventory"
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.StoredProcedure
                .Connection = conn
                .ExecuteNonQuery()
            End With
            conn.Close()
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function

    Public Function DeleteData(TableName As String) As String
        Dim sql As String = "TRUNCATE TABLE " & TableName
        Dim conn As New SqlConnection(ConnStr)
        conn.Open()
        Dim cmd As New SqlCommand
        With cmd
            .CommandText = sql
            .CommandType = CommandType.Text
            .Connection = conn
            .ExecuteNonQuery()
        End With
        conn.Close()
        Return ""
    End Function

    Public Function ConvertDate(strDate As String) As String
        Dim arr() As String = strDate.Split("/")
        Dim d As String = ""
        Dim m As String = ""
        Dim y As String = ""
        If arr.Length = 3 Then
            d = arr(0)
            m = arr(1)
            y = arr(2).Substring(0, 4)
            If CInt(y) > 2500 Then
                y = CInt(y) - 543
            End If
        End If

        Return "'" & y & "/" & m & "/" & d & "'"
    End Function

    Public Function ConvertDateTime(strDate As String) As String
        Dim arr_all() As String = strDate.Split(" ")
        Dim arr() As String = arr_all(0).Split("/")
        Dim d As String = ""
        Dim m As String = ""
        Dim y As String = ""
        If arr.Length = 3 Then
            d = arr(0)
            m = arr(1)
            y = arr(2).Substring(0, 4)
            If CInt(y) > 2500 Then
                y = CInt(y) - 543
            End If
        End If

        'Dim hh As String = ""
        'Dim mm As String = ""
        'Dim ss As String = ""
        'If arr_all.Length > 1 Then
        '    Dim arr_time() As String = arr_all(1).Split(":")
        '    If arr_time.Length = 3 Then
        '        hh = arr_time(0)
        '        mm = arr_time(1)
        '        ss = arr_time(2)
        '    End If

        'End If

        Dim time As String = "00:00:00"
        If arr_all.Length > 1 Then
            time = arr_all(1)
        End If


        Return "'" & y & "/" & m & "/" & d & " " & time & "'" '& hh & ":" & mm & ":" & ss & "'"
    End Function

    Public Function CheckExistsSP(StoreName As String) As String
        'sp_Initial_LUBE_Stock_Inventory

        Try
            Dim sql As String = "SELECT *  From sysobjects Where id = object_id(N'[dbo].[" & StoreName & "]')  And OBJECTPROPERTY(id, N'IsProcedure') = 1 "
            Dim da As New SqlDataAdapter(sql, ConnStr)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                sql = "DROP PROCEDURE " & StoreName & ""
                Dim conn As New SqlConnection(ConnStr)
                conn.Open()
                Dim cmd As New SqlCommand
                With cmd
                    .CommandText = sql
                    .CommandType = CommandType.Text
                    .Connection = conn
                    .ExecuteNonQuery()
                End With
                conn.Close()
            End If
            Return ""

        Catch ex As Exception
            Return "พบปัญหาในการนำเข้าข้อมูล :" & ex.ToString
        End Try

    End Function

End Class
