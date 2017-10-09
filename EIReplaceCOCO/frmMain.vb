Imports EIReplaceCOCO.Org.Mentalis.Files
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Public Class frmMain

    Dim ClsConn As New ConnectDatabase
    Dim ConnStr As String = ClsConn.ConnStr
    Dim ClsEncrypDecryp As New EncrypDecryp
    Dim ClsClobalFunction As New GlobalFunction

    Dim ClsImportData As New ImportData

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function FindWindowEx(ByVal parentHandle As IntPtr, ByVal childAfter As IntPtr, ByVal lclassName As String, ByVal windowTitle As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As Integer, ByVal lParam As String) As Integer
    End Function

#Region "Event"
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        writeLogResult("Start Program")
        Timer1.Interval = 100
        Timer1.Start()
        txtTransLog.TabStop = False
        ProgressBar1.Value = 0

        Try
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
        Catch ex As Exception
            MsgBox("ไม่สามารถเชื่อมต่อฐานข้อมูลได้", MsgBoxStyle.OkOnly)
            writeLogResult("Connection State is close.")
            Application.Exit()
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        lblDateTime.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
    End Sub

    Private Sub pbExit_Click(sender As Object, e As EventArgs) Handles pbExit.Click

        Using New Centered_MessageBox(Me)
            Dim confirm As DialogResult = MessageBox.Show("ต้องการปิดโปรแกรมใช่หรือไม่", "", MessageBoxButtons.OKCancel)
            If (confirm.Equals(DialogResult.OK)) Then
                writeLogResult("Exit Program")
                Application.Exit()
            End If
        End Using
    End Sub

    Private Sub pbExport_MouseHover(sender As Object, e As EventArgs) Handles pbExport.MouseHover
        Dim ToolTip1 As New ToolTip
        ToolTip1.SetToolTip(pbExport, "นำออกข้อมูล")
    End Sub

    Private Sub pbImport_MouseHover(sender As Object, e As EventArgs) Handles pbImport.MouseHover
        Dim ToolTip1 As New ToolTip
        ToolTip1.SetToolTip(pbImport, "นำเข้าข้อมูล")
    End Sub

    Private Sub pbExit_MouseHover(sender As Object, e As EventArgs) Handles pbExit.MouseHover
        Dim ToolTip1 As New ToolTip
        ToolTip1.SetToolTip(pbExit, "ปิด")
    End Sub


    Private Sub pbExport_Click(sender As Object, e As EventArgs) Handles pbExport.Click
        Dim strFileResule As String = Application.StartupPath & "\" & "_Result_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"
        Dim sw_rs As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(strFileResule, True)
        Try
            txtTransLog.Text = ""
            '#Save file Result, File Name _Result.txt  in Startup Path

            sw_rs.WriteLine(GetDateTime() & "เริ่มต้นนำออกข้อมูล")

            Dim frm As New frmConfirmPassword
            frm.Mode = 1
            If frm.ShowDialog(Me) = DialogResult.Cancel Then
                sw_rs.WriteLine(GetDateTime() & "Confirm Password : Cancel")
                sw_rs.Close()
                Exit Sub
            Else
                sw_rs.WriteLine(GetDateTime() & "Confirm Password : Success")
                txtTransLog.Text = GetDateTime() & "เริ่มต้นนำออกข้อมูล"
            End If

            EnableButton(False)
            ProgressBar1.Value = 3

            '#1. Connect Local Server Only
            '#2. Delete All File In Floder Export & Delete _Result.txt
            Dim path As String = Application.StartupPath & "\Export"
            If Directory.Exists(path) Then
                For Each _file As String In Directory.GetFiles(path)
                    File.Delete(_file)
                Next
            Else
                Directory.CreateDirectory(path)
            End If
            ProgressBar1.Value = 5


            '#BackupDB
            Application.DoEvents()
            sw_rs.WriteLine(GetDateTime() & "เริ่มต้น Backup Database")
            txtTransLog.Text = GetDateTime() & "เริ่มต้น Backup Database" & vbCrLf & txtTransLog.Text
            ProgressBar1.Value = 5
            Threading.Thread.Sleep(100)

            Application.DoEvents()
            Dim clsBKREDB As New BackUpRestoreDB
            Dim res As String = clsBKREDB.BackupDB(ProgressBar1)
            If res <> "" Then
                txtTransLog.Text = GetDateTime() & "พบปัญหาในการนำออกข้อมูล :" & res & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(GetDateTime() & "พบปัญหาในการนำออกข้อมูล :" & res)
            Else
                txtTransLog.Text = GetDateTime() & "Backup Database : Success" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(GetDateTime() & "Backup Database : Success")
            End If
            ProgressBar1.Value = 30
            Threading.Thread.Sleep(100)


            '#3. Export text file Name Same Table Name
            Dim arr() As String = {"APP_DATA", "TBPOS_PUMP_ALLOW", "TBMATERIAL", "TBBOM_USAGE", "TBMATTERIAL_SITE",
                "TBCONVERSION", "LKMAT_GROUP3", "LKDIVISION", "TBMAT_RECOMMEND", "APP_CONFIG", "POS_CONFIG", "TBUSER",
                "TSJOURNAL", "TSJOURNAL_DETAIL", "TSJOURNAL_PAYMENT", "TBPERIODS", "TBHOSE_HISTORY", "TBMATERIAL_HISTORY",
                "TBTANK_HISTORY", "TBPAY_IN", "TBPAYIN_PERIOD_LOG", "TBCARD", "TBMATERIAL_HISTORY_DESC", "TSSAFTDROP"}
            For Each item As String In arr
                Dim table_name As String = item
                Dim sql As String = ""
                Select Case table_name
                    Case "APP_DATA", "TBPOS_PUMP_ALLOW", "TBMATERIAL", "TBBOM_USAGE", "TBMATTERIAL_SITE", "TBCONVERSION", "LKMAT_GROUP3", "LKDIVISION",
                         "TBMAT_RECOMMEND", "APP_CONFIG", "POS_CONFIG", "TBUSER", "TSJOURNAL", "TSJOURNAL_DETAIL", "TSJOURNAL_PAYMENT", "TBPAYIN_PERIOD_LOG", "TBCARD", "TBMATERIAL_HISTORY_DESC", "TSSAFTDROP"
                        sql &= "SELECT * FROM " & table_name
                    Case "TBPAY_IN"
                        sql = "SELECT [PAYIN_ID] "
                        sql &= ",[REFBILL_NO]"
                        sql &= ",[TRANSFER_DATE]"
                        sql &= ",[TRAN_DATE]"
                        sql &= ",[BUS_DATE]"
                        sql &= ",REPLACE(REPLACE([SHIFT_DESCRIPTION], Char(13), '$$'), CHAR(10), '&&') as SHIFT_DESCRIPTION"
                        sql &= ",[LAST_CLOSE_SHIFT_DT]"
                        sql &= ",[FILEPATH]"
                        sql &= ",[FILENAME]"
                        sql &= ",[TYPE]"
                        sql &= ",[PAYMENT_TYPE]"
                        sql &= ",[AMOUNTREC]"
                        sql &= ",[AMOUNT]"
                        sql &= ",[AMOUNT_DIFF]"
                        sql &= ",REPLACE(REPLACE([REMARK], Char(13), '$$'), CHAR(10), '&&') REMARK "
                        sql &= ",[STATUS_SAP]"
                        sql &= ",[STATUS]"
                        sql &= ",[NO_SALE_STATUS]"
                        sql &= ",[CREATEDATE]"
                        sql &= ",[CREATEBY]"
                        sql &= ",[UPDATEDATE]"
                        sql &= ",[UPDATEBY]"
                        sql &= " FROM " & table_name
                    Case "TBPERIODS"
                        sql &= "SELECT * FROM TBPERIODS WHERE DAY_ID IN (SELECT DISTINCT  DAY_ID FROM TSJOURNAL)"

                    Case "TBHOSE_HISTORY"
                        sql &= "SELECT * FROM TBHOSE_HISTORY WHERE PERIOD_ID IN (SELECT PERIOD_ID FROM TBPERIODS WHERE DAY_ID IN (SELECT TOP 1 DAY_ID FROM TSJOURNAL))"

                    Case "TBMATERIAL_HISTORY"
                        sql &= "SELECT * FROM TBMATERIAL_HISTORY  WHERE PERIOD_ID IN (SELECT PERIOD_ID FROM TBPERIODS WHERE DAY_ID IN (SELECT TOP 1 DAY_ID FROM TSJOURNAL))"

                    Case "TBTANK_HISTORY"
                        sql &= "SELECT * FROM TBTANK_HISTORY  WHERE PERIOD_ID IN (SELECT PERIOD_ID FROM TBPERIODS WHERE DAY_ID IN (SELECT TOP 1 DAY_ID FROM TSJOURNAL))"
                End Select

                Dim da As New SqlDataAdapter(sql, ConnStr)
                Dim dt As New DataTable
                da.Fill(dt)

                '#4. Save text file in floder Application.StartupPath & Export
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then

                    If table_name = "TBUSER" Then
                        For j As Integer = 0 To dt.Rows.Count - 1
                            dt.Rows(j)("PASSWORD") = ClsEncrypDecryp.Decrypt(dt.Rows(j)("USERNAME").ToString, dt.Rows(j)("PASSWORD").ToString)
                        Next
                    End If

                    Dim strFile As String = path & "\" & table_name & ".txt"

                    'Add column.
                    Dim strcolumn As String = ""
                    For Each column As DataColumn In dt.Columns
                        strcolumn += "|" & column.ColumnName
                    Next
                    strcolumn = strcolumn.Substring(1)

                    'Add data.
                    Dim result As New StringBuilder
                    Dim i As Integer = 0
                    For Each row As DataRow In dt.Rows
                        Dim line As String = ""
                        For Each column As DataColumn In dt.Columns
                            line += "|" & row(column.ColumnName)
                        Next
                        result.AppendLine(line.Substring(1))
                        i += 1
                    Next
                    Using sw As StreamWriter = New StreamWriter(strFile)
                        sw.WriteLine(strcolumn)
                        sw.WriteLine(result.ToString)
                        sw.Close()
                    End Using

                    '#Save file Result, File Name _Result.txt  in Startup Path
                    sw_rs.WriteLine(GetDateTime() & table_name & "  (" & dt.Rows.Count.ToString & " Row)")

                    Application.DoEvents()
                    txtTransLog.Text = GetDateTime() & table_name & "  (" & dt.Rows.Count.ToString & " Row)" & vbCrLf & txtTransLog.Text
                    Threading.Thread.Sleep(100)
                End If
                If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2
            Next ' end for arr


            '#5. Copy File
            Dim msgcopyret As String = ""
            Dim ClsCopyAndReplaceFile As New CopyAndReplaceFile

            '--5.1 Copy CloseDay File
            Dim ret As String = ClsCopyAndReplaceFile.CopyCloseDayFile
            If ret = "" Then
                msgcopyret = "คัดลอกไฟล์ปิดวันสำเร็จ"
            Else
                msgcopyret = ret
            End If
            Application.DoEvents()
            sw_rs.WriteLine(GetDateTime() & msgcopyret)
            txtTransLog.Text = GetDateTime() & msgcopyret & vbCrLf & txtTransLog.Text
            If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2

            '--5.2 Copy POSDB File
            ret = ClsCopyAndReplaceFile.CopyPOSDBFile
            If ret = "" Then
                msgcopyret = "คัดลอกไฟล์ POSDB สำเร็จ"
            Else
                msgcopyret = ret
            End If
            Application.DoEvents()
            sw_rs.WriteLine(GetDateTime() & msgcopyret)
            txtTransLog.Text = GetDateTime() & msgcopyret & vbCrLf & txtTransLog.Text
            If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2

            '--5.3 Copy Application Log File
            ret = ClsCopyAndReplaceFile.CopyApplicationLogFile
            If ret = "" Then
                msgcopyret = "คัดลอกไฟล์ Application Log สำเร็จ"
            Else
                msgcopyret = ret
            End If
            Application.DoEvents()
            sw_rs.WriteLine(GetDateTime() & msgcopyret)
            txtTransLog.Text = GetDateTime() & msgcopyret & vbCrLf & txtTransLog.Text
            If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2


            txtTransLog.Text = GetDateTime() & "สิ้นสุดการนำออกข้อมูล" & vbCrLf & txtTransLog.Text

            '#6. Final Export Save file Result, File Name _Result.txt  in Startup Path
            sw_rs.WriteLine(GetDateTime() & "สิ้นสุดการนำออกข้อมูล")
            sw_rs.Close()

            EnableButton(True)
            ProgressBar1.Value = 100
        Catch ex As Exception
            sw_rs.WriteLine(GetDateTime() & "Export Fail")
            sw_rs.Close()

            Using New Centered_MessageBox(Me)
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK)
            End Using

        End Try
    End Sub

    Private Sub pbImport_Click(sender As Object, e As EventArgs) Handles pbImport.Click
        Dim strFileResule As String = Application.StartupPath & "\" & "_Result_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"
        Dim sw_rs As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(strFileResule, True)

        Try
            txtTransLog.Text = ""
            '#Save file Result, File Name _Result.txt  in Startup Path
            sw_rs.WriteLine(GetDateTime() & "เริ่มต้นนำเข้าข้อมูล")

            Dim frm As New frmConfirmPassword
            frm.Mode = 2
            If frm.ShowDialog(Me) = DialogResult.Cancel Then
                sw_rs.WriteLine(GetDateTime() & "Confirm Password : Cancel")
                sw_rs.Close()
                Exit Sub
            Else
                sw_rs.WriteLine(GetDateTime() & "Confirm Password : Success")
                txtTransLog.Text = GetDateTime() & "เริ่มต้นนำเข้าข้อมูล"
            End If

            EnableButton(False)
            ProgressBar1.Value = 3

            '#RestoreDB
            Application.DoEvents()
            sw_rs.WriteLine(GetDateTime() & "เริ่มต้น Restore Database")
            txtTransLog.Text = GetDateTime() & "เริ่มต้น Restore Database" & vbCrLf & txtTransLog.Text
            ProgressBar1.Value = 5
            Threading.Thread.Sleep(100)

            Application.DoEvents()
            Dim clsBKREDB As New BackUpRestoreDB
            Dim res As String = clsBKREDB.RestoreDB(ProgressBar1)
            If res <> "" Then
                txtTransLog.Text = GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล :" & res & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล :" & res)
            Else
                txtTransLog.Text = GetDateTime() & "Restore Database : Success" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(GetDateTime() & "Restore Database : Success")
            End If
            ProgressBar1.Value = 30
            Threading.Thread.Sleep(100)


            Dim dt_appconfig As New DataTable
            Dim path As String = Application.StartupPath & "\Export"
            If Directory.Exists(path) Then
                For Each _file As String In Directory.GetFiles(path)
                    Dim file_name As String = System.IO.Path.GetFileName(_file)
                    Dim str() As String = file_name.Split(".")
                    Dim table_name As String = ""
                    If str.Length > 0 Then
                        table_name = str(0)
                    End If

                    'Delete Data
                    ClsClobalFunction.DeleteData(table_name)

                    'Insert Data
                    Dim dt As New DataTable
                    Dim dr As DataRow
                    Dim i As Integer = 0
                    Dim sr As StreamReader = New StreamReader(_file)
                    Dim line As String = ""
                    Do While sr.Peek() >= 0
                        line = sr.ReadLine()
                        If i = 0 AndAlso line <> "" Then
                            Dim strColumn() As String = line.Split("|")
                            For j As Integer = 0 To strColumn.Length - 1
                                dt.Columns.Add(strColumn(j))
                            Next
                        Else
                            Dim strValue() As String = line.Split("|")
                            dr = dt.NewRow
                            If strValue.Length = dt.Columns.Count Then
                                For j As Integer = 0 To strValue.Length - 1
                                    dr(j) = strValue(j)
                                Next
                                If dr(0).ToString <> "" Then
                                    dt.Rows.Add(dr)
                                End If
                            End If
                        End If

                        i = i + 1
                    Loop
                    sr.Close()

                    Dim result As String = ""
                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                        result = ClsImportData.InsertData(table_name, dt)
                    End If

                    If table_name.ToUpper = "APP_CONFIG" Then
                        dt_appconfig = dt
                    End If


                    If result <> "" Then
                        txtTransLog.Text = GetDateTime() & table_name & "ไม่สามารถนำเข้าข้อมูลได้ !" & vbCrLf & result & vbCrLf & txtTransLog.Text
                        sw_rs.WriteLine(GetDateTime() & table_name & "ไม่สามารถนำเข้าข้อมูลได้ !" & vbCrLf & result)
                        sw_rs.Close()
                        Exit Sub
                    End If


                    If table_name.ToUpper <> "APP_CONFIG" Then
                        'Save Transection Log
                        Application.DoEvents()
                        txtTransLog.Text = GetDateTime() & table_name & "  (" & dt.Rows.Count.ToString & " Row)" & vbCrLf & txtTransLog.Text
                        Threading.Thread.Sleep(100)

                        'Save Result Log
                        sw_rs.WriteLine(GetDateTime() & table_name & "  (" & dt.Rows.Count.ToString & " Row)")
                    End If
                    If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2
                Next

                'Create sp_Import_Product_To_Inventory
                Dim retsp As String = ""
                Dim retdropsp As String = ClsClobalFunction.CheckExistsSP("sp_Import_Product_To_Inventory")
                If retdropsp = "" Then
                    retsp = CreateStoreImportPDToInvent()
                    If retsp = "" Then
                        Application.DoEvents()
                        txtTransLog.Text = GetDateTime() & "Create sp_Import_Product_To_Inventory" & vbCrLf & txtTransLog.Text
                        sw_rs.WriteLine(GetDateTime() & "Create sp_Import_Product_To_Inventory")
                        Threading.Thread.Sleep(100)
                    End If
                Else
                    txtTransLog.Text = GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล : Drop sp_Import_Product_To_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล :  Drop sp_Import_Product_To_Inventory" & retdropsp)
                End If

                If retsp <> "" Then
                    Application.DoEvents()
                    txtTransLog.Text = GetDateTime() & "Cant Create sp_Import_Product_To_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(GetDateTime() & "Cant Create sp_Import_Product_To_Inventory" & retsp)
                    Threading.Thread.Sleep(100)
                End If

                'call sp sp_Import_Product_To_Inventory
                Application.DoEvents()
                If ClsClobalFunction.CallSPImportProduct() = False Then
                    txtTransLog.Text = GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล : Call sp_Import_Product_To_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล : Call sp_Import_Product_To_Inventory")
                Else
                    txtTransLog.Text = GetDateTime() & "Call sp_Import_Product_To_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(GetDateTime() & "Call sp_Import_Product_To_Inventory")
                End If
                If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2

                'Update ISSHOWINPOS =0 , ISRECOMMEND=0
                Application.DoEvents()
                Dim ret0 As Integer = ClsClobalFunction.UpdateDefaultProduct()
                txtTransLog.Text = GetDateTime() & "Update ISSHOWINPOS = 0,ISSHOWINPOS =0 (" & ret0 & " Row)" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(GetDateTime() & "Update ISSHOWINPOS = 0,ISSHOWINPOS =0 (" & ret0 & " Row)")
                Threading.Thread.Sleep(100)


                'Update ข้อมูลที่ POSDB.dbo.PRODUCTS.ISSHOWINPOS ให้เป็น 1 เฉพาะรหัสผลิตภัณฑ์ที่มีอยู่ในตาราง POSDB.dbo.TBMATERIAL_SITE
                Application.DoEvents()
                Dim ret1 As Integer = ClsClobalFunction.UpdateISSHOWINPOS()
                txtTransLog.Text = GetDateTime() & "Update ISSHOWINPOS (" & ret1 & " Row)" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(GetDateTime() & "Update ISSHOWINPOS (" & ret1 & " Row)")
                Threading.Thread.Sleep(100)


                'Update ข้อมูลที่ POSDB.dbo.PRODUCTS.ISRECOMMEND ให้เป็น 1 เฉพาะรหัสผลิตภัณฑ์ที่มีอยู่ในตาราง POSDB.dbo.TBMAT_RECOMMENED
                Application.DoEvents()
                Dim ret2 As Integer = ClsClobalFunction.UpdateISRECOMMEND()
                txtTransLog.Text = GetDateTime() & "Update ISRECOMMEND (" & ret2 & " Row)" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(GetDateTime() & "Update ISRECOMMEND (" & ret2 & " Row)")
                Threading.Thread.Sleep(100)
                If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2

                'Create sp_Initial_LUBE_Stock_Inventory
                retsp = ""
                retdropsp = ClsClobalFunction.CheckExistsSP("sp_Initial_LUBE_Stock_Inventory")
                If retdropsp = "" Then
                    retsp = CreateStoreInitialLUBE()
                    If retsp = "" Then
                        Application.DoEvents()
                        txtTransLog.Text = GetDateTime() & "Create sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                        sw_rs.WriteLine(GetDateTime() & "Create sp_Initial_LUBE_Stock_Inventory")
                        Threading.Thread.Sleep(100)
                    End If
                Else
                    txtTransLog.Text = GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล : Drop sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล :  Drop sp_Initial_LUBE_Stock_Inventory" & retdropsp)
                End If

                If retsp <> "" Then
                    Application.DoEvents()
                    txtTransLog.Text = GetDateTime() & "Cant Create sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(GetDateTime() & "Cant Create sp_Initial_LUBE_Stock_Inventory" & retsp)
                    Threading.Thread.Sleep(100)
                End If


                'call sp_Initial_LUBE_Stock_Inventory
                Application.DoEvents()
                Dim ret_CallSPInitialLUBE As String = ClsClobalFunction.CallSPInitialLUBE()
                If ret_CallSPInitialLUBE <> "" Then
                    txtTransLog.Text = GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล : Call sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล : Call sp_Initial_LUBE_Stock_Inventory : " & ret_CallSPInitialLUBE)
                Else
                    txtTransLog.Text = GetDateTime() & "Call sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(GetDateTime() & "Call sp_Initial_LUBE_Stock_Inventory")
                End If
                If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2

                'RunScriptSQL
                Dim script_path As String = Application.StartupPath & "\Scripts"
                If Directory.Exists(script_path) Then
                    Dim di As New DirectoryInfo(script_path)
                    Dim files As FileSystemInfo() = di.GetFileSystemInfos()
                    Dim orderedFiles As Array = files.OrderBy(Function(f) f.FullName).ToArray
                    For i As Integer = 0 To orderedFiles.Length - 1
                        Dim _file As String = DirectCast(orderedFiles, System.IO.FileSystemInfo())(i).FullName
                        Dim _file_name As String = DirectCast(orderedFiles, System.IO.FileSystemInfo())(i).Name

                        If _file_name.ToLower <> "2_sp_Initial_LUBE_Stock_Inventory.sql".ToLower And _file_name.ToLower <> "1_sp_Import_Product_To_Inventory.sql".ToLower Then
                            Dim ret_RunScriptSQL As String = RunScriptSQL(_file)
                            If ret_RunScriptSQL = "" Then
                                Application.DoEvents()
                                txtTransLog.Text = GetDateTime() & "Call " & _file_name & vbCrLf & txtTransLog.Text
                                sw_rs.WriteLine(GetDateTime() & "Call " & _file_name)
                                Threading.Thread.Sleep(100)
                            Else
                                Application.DoEvents()
                                txtTransLog.Text = GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล : Call " & _file_name & vbCrLf & txtTransLog.Text
                                sw_rs.WriteLine(GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล : Call " & _file_name & "      " & ret_RunScriptSQL)
                                Threading.Thread.Sleep(100)
                            End If
                        End If
                        If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2
                    Next
                End If


                'Update APP_Config
                If Not dt_appconfig Is Nothing AndAlso dt_appconfig.Rows.Count > 0 Then
                    Dim cnt As Integer = 0
                    For i As Integer = 0 To dt_appconfig.Rows.Count - 1
                        Dim config_key As String = dt_appconfig.Rows(i)("CONFIG_KEY").ToString
                        Dim config_value As String = dt_appconfig.Rows(i)("CONFIG_VALUE").ToString
                        If ClsClobalFunction.Update_APP_Config(config_key, config_value) > 0 Then
                            cnt += 1
                        End If
                    Next
                    Application.DoEvents()
                    txtTransLog.Text = GetDateTime() & "Update APP_CONFIG" & "  (" & cnt & " Row)" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(GetDateTime() & "Update APP_CONFIG" & "  (" & cnt & " Row)")
                    Threading.Thread.Sleep(100)
                End If
                If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2


                'Copy File
                Dim msgcopyret As String = ""
                Dim ClsCopyAndReplaceFile As New CopyAndReplaceFile

                '--Copy CloseDay File
                Dim ret As String = ClsCopyAndReplaceFile.ReplaceCloseDayFile
                If ret = "" Then
                    msgcopyret = "คัดลอกไฟล์ปิดวันสำเร็จ"
                Else
                    msgcopyret = ret
                End If
                Application.DoEvents()
                sw_rs.WriteLine(GetDateTime() & msgcopyret)
                txtTransLog.Text = GetDateTime() & msgcopyret & vbCrLf & txtTransLog.Text
                If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2

                '--Copy POSDB File
                ret = ClsCopyAndReplaceFile.ReplacePOSDBFile
                If ret = "" Then
                    msgcopyret = "คัดลอกไฟล์ POSDB สำเร็จ"
                Else
                    msgcopyret = ret
                End If
                Application.DoEvents()
                sw_rs.WriteLine(GetDateTime() & msgcopyret)
                txtTransLog.Text = GetDateTime() & msgcopyret & vbCrLf & txtTransLog.Text
                If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2

                '--Copy Application Log File
                ret = ClsCopyAndReplaceFile.ReplaceApplicationLogFile
                If ret = "" Then
                    msgcopyret = "คัดลอกไฟล์ Application Log สำเร็จ"
                Else
                    msgcopyret = ret
                End If
                Application.DoEvents()
                sw_rs.WriteLine(GetDateTime() & msgcopyret)
                txtTransLog.Text = GetDateTime() & msgcopyret & vbCrLf & txtTransLog.Text
                If ProgressBar1.Value >= 90 Then ProgressBar1.Value = 90 Else ProgressBar1.Value = ProgressBar1.Value + 2

                txtTransLog.Text = GetDateTime() & "สิ้นสุดการนำเข้าข้อมูล" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(GetDateTime() & "สิ้นสุดการนำเข้าข้อมูล")

                EnableButton(True)
                ProgressBar1.Value = 100
            Else
                Using New Centered_MessageBox(Me)
                    MessageBox.Show("ไม่พบรายการสำหรับนำเข้าข้อมูล", "", MessageBoxButtons.OK)
                End Using
                EnableButton(True)
            End If

        Catch ex As Exception
            Application.DoEvents()
            txtTransLog.Text = GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล " & ex.ToString & vbCrLf & txtTransLog.Text
            sw_rs.WriteLine(GetDateTime() & "พบปัญหาในการนำเข้าข้อมูล " & ex.ToString)
            Threading.Thread.Sleep(100)


            Using New Centered_MessageBox(Me)
                MessageBox.Show("ไม่สามารถนำเข้าข้อมูลได้ !" & vbCrLf & ex.ToString, "", MessageBoxButtons.OK)
            End Using
        End Try
        sw_rs.Close()
    End Sub

#End Region

#Region "Sub&Function"

    Function GetDateTime()
        Return ClsClobalFunction.GetDateTime()
    End Function

    Sub EnableButton(IsEnable As Boolean)

        pbExport.Enabled = IsEnable
        pbImport.Enabled = IsEnable
        If IsEnable Then
            pbExport.BackgroundImage = My.Resources.export_th__2_
            pbImport.BackgroundImage = My.Resources.Import_th
        Else
            pbExport.BackgroundImage = My.Resources.export_th_dis
            pbImport.BackgroundImage = My.Resources.import_th_dis
        End If

    End Sub

    Sub writeLogResult(strResult As String)
        Try
            Dim strFileResule As String = Application.StartupPath & "\" & "_Result_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"
            Dim strdate As String = DateTime.Now.ToString("yyyyMMdd hh:MM:ss") & "  "
            Dim sw_rs As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(strFileResule, True)
            sw_rs.WriteLine(strdate & strResult)
            sw_rs.Close()
        Catch ex As Exception

        End Try

    End Sub

    Function CreateStoreInitialLUBE() As String
        'True = Success , False = Fail
        Try

            Dim path As String = Application.StartupPath & "\" & "Scripts\2_sp_Initial_LUBE_Stock_Inventory.sql"
            Dim ret As String = RunScriptSQL(path)
            Return ret

        Catch ex As Exception
            Return ex.ToString
        End Try


    End Function

    Function CreateStoreImportPDToInvent() As String
        'True = Success , False = Fail
        Try

            Dim path As String = Application.StartupPath & "\" & "Scripts\1_sp_Import_Product_To_Inventory.sql"
            Dim ret As String = RunScriptSQL(path)
            Return ret

        Catch ex As Exception
            Return ex.ToString
        End Try


    End Function


#End Region

#Region "RunScript"
    Function RunScriptSQL(path As String) As String
        Dim lpcstatus_str As String = ""
        lpcstatus_str = Me.ExecScriptFile(path)

        Return lpcstatus_str
    End Function

    Function ExecScriptFile(ByVal pscript_file As String) As String
        Dim lresult_str As String = ""
        Try
            RunCommandCom("Start /min notepad """ & pscript_file & """", "", False)
            Dim fileName As String() = pscript_file.Split("\")
            Dim script = ReadTextFromNotePad(fileName(fileName.Length - 1), 2000)
            RunCommandCom("Taskkill /IM notepad.exe", "", False)

            script = Regex.Replace(script, "/\*(.|\n)*?\*/", "")
            Dim commandStrings As IEnumerable(Of String) = Regex.Split(script, "^\s*GO\s*$|^\s*GO", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
            For Each cmd As String In commandStrings
                If (cmd.Trim() <> "") Then
                    lresult_str = Me.ExecNoneQuery(cmd)
                    If (lresult_str <> "") Then
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            lresult_str = ex.ToString
        End Try

        Return lresult_str
    End Function

    Function ExecNoneQuery(ByVal psql_str As String) As String
        Dim lresult_str As String = ""
        Dim lcomm As SqlClient.SqlCommand = Nothing
        Try

            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = psql_str
                .CommandType = CommandType.Text
                .CommandTimeout = 120
                .Connection = conn
                .ExecuteNonQuery()
            End With
            conn.Close()
            lresult_str = ""

        Catch ex As Exception
            lresult_str = ex.ToString
        End Try
        Return lresult_str
    End Function


    Const WM_SETTEXT As Integer = &HC
    Const WM_GETTEXT As Integer = &HD
    Const WM_GETTEXTLENGTH As Integer = &HE
    Function ReadTextFromNotePad(fileName As String, timeOut As Integer) As String
        Dim result As String = "Time Out"
        For time = 0 To timeOut
            Dim hParent As IntPtr = FindWindowEx(IntPtr.Zero, hParent, "Notepad", fileName & " - Notepad")
            If Not hParent.Equals(IntPtr.Zero) Then
                Dim hChild As IntPtr = FindWindowEx(hParent, hChild, "Edit", vbNullString)
                If Not hChild.Equals(IntPtr.Zero) Then
                    Dim txtlen As Integer = SendMessage(hChild, WM_GETTEXTLENGTH, 0, vbNullString)
                    Dim txt As String = Space(txtlen + 1)
                    SendMessage(hChild, WM_GETTEXT, txtlen + 1, txt)
                    Return txt
                Else
                    result = "Child Window Not Found"
                End If
            Else
                result = "Main Window Not Found"
            End If
            System.Threading.Thread.Sleep(1)
        Next
        Return result
    End Function

    Sub RunCommandCom(command As String, arguments As String, permanent As Boolean)
        Dim p As Process = New Process()
        Dim pi As ProcessStartInfo = New ProcessStartInfo()
        pi.Arguments = " " + If(permanent = True, "/K", "/C") + " " + command + " " + arguments
        pi.FileName = "cmd.exe"
        pi.CreateNoWindow = True
        pi.WindowStyle = ProcessWindowStyle.Hidden
        p.StartInfo = pi
        p.Start()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ClsCopyAndReplaceFile As New CopyAndReplaceFile
        Dim ret As String = ClsCopyAndReplaceFile.ReplaceApplicationLogFile
        If ret = "" Then
            MessageBox.Show("True", "", MessageBoxButtons.OK)
        Else
            MessageBox.Show("False" & ret, "", MessageBoxButtons.OK)
        End If

    End Sub
#End Region


End Class
