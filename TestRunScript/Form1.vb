Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Text

Public Class Form1

    Dim password As String = "pTT!CT01"
    'Dim password As String = "1qaz@WSX"
    Dim ConnStr As String = "Data Source=.;Initial Catalog=POSDB;User ID=sa;Password=" & password & ";Connect Timeout=1;"
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function FindWindowEx(ByVal parentHandle As IntPtr, ByVal childAfter As IntPtr, ByVal lclassName As String, ByVal windowTitle As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As Integer, ByVal lParam As String) As Integer
    End Function

    Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim flltax_file As String = Application.StartupPath & "\" & "Script.sql"
        Dim ret_RunScriptFullTax As String = RunScriptSQL(flltax_file, ConnStr)
        'Dim ret_RunScriptFullTax As String = RunScriptSQL3()
        If ret_RunScriptFullTax = "" Then
            txtTransLog.Text = "Succcess!"
        Else
            txtTransLog.Text = ret_RunScriptFullTax
        End If

    End Sub

#Region "RunScript"
    'sqlcmd -U sa -P 1qaz@WSX -S (local) -d POSDB -i D:\TestRunScript\Script.sql

    Function RunScriptSQL3() As String
        Try
            'RunCommandCom("sqlcmd -U sa -P " & password & " -S (local) -d POSDB -i D:\TestRunScript\Script.sql", "/W", True)
            RunCommandCom("sqlcmd -U sa -P " & password & " -S (local) -d POSDB -i D:\TestRunScript\Script.sql -o D:\TestRunScript\cmdResult.txt", "/W", True)
            'RunCommandCom("sqlcmd -U sa -P " & password & " -S (local) -d POSDB -Q ""Select Top 2 *,'xx' as xx from App_Config"" -o D:\TestRunScript\cmdResult.txt", "/W", True)

            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function

    Function RunScriptSQL2() As String
        Try
            Dim file As New FileInfo(Application.StartupPath & "\" & "Script.sql")
            Dim script As String = file.OpenText().ReadToEnd()

            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = script
                .CommandType = CommandType.Text
                .CommandTimeout = 120
                .Connection = conn
                .ExecuteNonQuery()
            End With
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function

    Function RunScriptSQL(path As String, conn As String) As String
        Dim lpcstatus_str As String = ""
        lpcstatus_str = Me.ExecScriptFile(path, conn)

        Return lpcstatus_str
    End Function

    Function ExecScriptFile(ByVal pscript_file As String, conn As String) As String
        Dim lresult_str As String = ""
        Dim tempcmd As String = ""
        'Dim strFileResule As String = Application.StartupPath & "\" & "TestResult.txt"
        'Dim sw_rs As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(strFileResule, True)
        Try
            RunCommandCom("Start /min notepad """ & pscript_file & """", "", False)
            Dim fileName As String() = pscript_file.Split("\")

            'txtResult.Text &= "pscript_file" & vbCrLf
            'txtResult.Text &= pscript_file & vbCrLf

            'For Each item As String In fileName
            '    txtResult.Text &= "fileName" & vbCrLf
            '    txtResult.Text &= item & vbCrLf
            'Next

            Dim script = ReadTextFromNotePad(fileName(fileName.Length - 1), 2000)
            RunCommandCom("Taskkill /IM notepad.exe", "", False)

            txtResult.Text &= "Script1" & vbCrLf
            txtResult.Text &= script & vbCrLf


            If script.ToLower = "main window not found" Then
                script = ReadTextFromNotePad(fileName(fileName.Length - 1).Replace(".sql", ""), 2000)
            End If

            txtResult.Text &= "Script2" & vbCrLf
            txtResult.Text &= script & vbCrLf

            script = Regex.Replace(script, "/\*(.|\n)*?\*/", "")
            Dim commandStrings As IEnumerable(Of String) = Regex.Split(script, "^\s*GO\s*$|^\s*GO", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
            For Each cmd As String In commandStrings
                txtResult.Text &= cmd
                If (cmd.Trim() <> "") Then
                    lresult_str = Me.ExecNoneQuery(cmd, conn)
                    If (lresult_str <> "") Then
                        Exit For
                    End If
                End If

                tempcmd = cmd
                'sw_rs.WriteLine(cmd)
            Next


        Catch ex As Exception
            lresult_str = "ExecScriptFile" & tempcmd & vbCrLf & vbCrLf & ex.ToString
        End Try
        'sw_rs.Close()
        Return lresult_str
    End Function

    Function ExecNoneQuery(ByVal psql_str As String, connS As String) As String
        Dim lresult_str As String = ""
        Dim lcomm As SqlClient.SqlCommand = Nothing
        Try

            Dim conn As New SqlConnection(connS)
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
            lresult_str = "ExecNoneQuery" & psql_str & vbCrLf & vbCrLf & ex.ToString
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

#End Region
End Class
