Imports System.IO
Imports ICSharpCode.SharpZipLib.Core
Imports ICSharpCode.SharpZipLib.Zip

Public Class CopyAndReplaceFile

    Public Function CopyCloseDayFile() As String
        'BO:Copy zip file ปิดวัน ณ วันที่เข้า Replace
        'E:\POSSible2\POSBACK\BACKUP\ZipFile\ -->> Date Modified lastest 30Days

        Try
            'Get lastest 30 File
            Dim path As String = "E:\POSSible2\POSBACK\BACKUP\ZipFile"
            Dim di As New DirectoryInfo(path)
            Dim files As FileSystemInfo() = di.GetFileSystemInfos()
            Dim orderedFiles As Array = files.OrderBy(Function(f) f.CreationTime).ToArray

            'CreateDirectory
            Dim destinationPath As String = Application.StartupPath & "\ZipFile"
            If Directory.Exists(destinationPath) = False Then
                Directory.CreateDirectory(destinationPath)
            Else
                For Each _file As String In Directory.GetFiles(destinationPath)
                    File.Delete(_file)
                Next
            End If

            'Copy File
            For i As Integer = orderedFiles.Length - 1 To IIf(orderedFiles.Length > 30, orderedFiles.Length - 30, 0) Step -1
                Dim _file As String = DirectCast(orderedFiles, System.IO.FileSystemInfo())(i).FullName
                Dim _filename As String = DirectCast(orderedFiles, System.IO.FileSystemInfo())(i).Name
                File.Copy(_file, IIf(destinationPath.EndsWith("\"), destinationPath, destinationPath & "\") & _filename)
            Next
            Return ""
        Catch ex As Exception
            Return "พบปัญหาในการคัดลอกไฟล์ : ZipFile :" & ex.ToString
        End Try
    End Function

    Public Function CopyPOSDBFile() As String
        'BO:Zip file Backup Database เฉพาะวันที่ 01 ของทุกเดือน   

        Try
            'CreateDirectory
            Dim destinationPath As String = Application.StartupPath & "\BackUpDB"
            If Directory.Exists(destinationPath) = False Then
                Directory.CreateDirectory(destinationPath)
            Else
                For Each _file As String In Directory.GetFiles(destinationPath)
                    File.Delete(_file)
                Next
            End If

            'Get All File
            Dim path As String = "D:\BackUpDB"
            Dim di As New DirectoryInfo(path)
            Dim files As FileSystemInfo() = di.GetFileSystemInfos()

            'Copy File
            For Each _file As FileSystemInfo In files
                Dim _filename As String = _file.Name
                Dim _last_date As String = _file.LastWriteTime.ToString("ddMMyyyy", New System.Globalization.CultureInfo("en-US"))
                If _last_date.Substring(0, 2) = "01" Then
                    File.Copy(_file.FullName, IIf(destinationPath.EndsWith("\"), destinationPath, destinationPath & "\") & _filename)
                End If
            Next

            Return ""
        Catch ex As Exception
            Return "พบปัญหาในการคัดลอกไฟล์ : BackUpDB :" & ex.ToString
        End Try
    End Function

    Public Function CopyApplicationLogFile() As String
        'POS:Copy Application log มาไว้เครื่อง POS ใหม่ด้วย เพื่อตรวจสอบ Log ย้อนหลัง ให้ zip

        Try
            'CreateDirectory
            Dim destinationPath As String = Application.StartupPath & "\POSLogs"
            If Directory.Exists(destinationPath) = False Then
                Directory.CreateDirectory(destinationPath)
            Else
                For Each _file As String In Directory.GetFiles(destinationPath)
                    File.Delete(_file)
                Next
            End If

            Dim zipPath As String = "C:\POSLogs"
            Dim zipToPathFile As String = destinationPath & "\POSLogs.zip"
            Dim Passw As String = ""

            'Get All File
            Dim files As String() = Directory.GetFiles(zipPath)

            'Zip File
            CreateZIPFile(zipToPathFile, Passw, files)

            Return ""
        Catch ex As Exception
            Return "พบปัญหาในการคัดลอกไฟล์ : Application log :" & ex.ToString
        End Try

    End Function


    Public Function ReplaceCloseDayFile() As String
        Try
            'Get All File
            Dim path As String = Application.StartupPath & "\ZipFile"
            Dim di As New DirectoryInfo(path)
            Dim files As FileSystemInfo() = di.GetFileSystemInfos()

            'CreateDirectory
            Dim destinationPath As String = "E:\POSSible2\POSBACK\BACKUP\ZipFile"
            If Directory.Exists(destinationPath) = False Then
                Directory.CreateDirectory(destinationPath)
            End If

            'Copy File
            For Each _file As FileSystemInfo In files
                Dim _filename As String = _file.Name
                File.Copy(_file.FullName, IIf(destinationPath.EndsWith("\"), destinationPath, destinationPath & "\") & _filename, True)
            Next
            Return ""
        Catch ex As Exception
            Return "พบปัญหาในการคัดลอกไฟล์ : ZipFile :" & ex.ToString
        End Try
    End Function

    Public Function ReplacePOSDBFile() As String
        Try
            'CreateDirectory
            Dim destinationPath As String = "D:\BackUpDB"
            If Directory.Exists(destinationPath) = False Then
                Directory.CreateDirectory(destinationPath)
            End If

            'Get All File
            Dim path As String = Application.StartupPath & "\BackUpDB"
            Dim di As New DirectoryInfo(path)
            Dim files As FileSystemInfo() = di.GetFileSystemInfos()

            'Copy File
            For Each _file As FileSystemInfo In files
                Dim _filename As String = _file.Name
                File.Copy(_file.FullName, IIf(destinationPath.EndsWith("\"), destinationPath, destinationPath & "\") & _filename, True)
            Next

            Return ""
        Catch ex As Exception
            Return "พบปัญหาในการคัดลอกไฟล์ : BackUpDB :" & ex.ToString
        End Try
    End Function

    Public Function ReplaceApplicationLogFile() As String
        Try

            'CreateDirectory
            Dim destinationPath As String = "C:\POSLogs"
            If Directory.Exists(destinationPath) = False Then
                Directory.CreateDirectory(destinationPath)
            End If

            'Get All File
            Dim path As String = Application.StartupPath & "\POSLogs"
            Dim di As New DirectoryInfo(path)
            Dim files As FileSystemInfo() = di.GetFileSystemInfos()

            'Copy File
            For Each _file As FileSystemInfo In files
                Dim _filename As String = _file.Name
                File.Copy(_file.FullName, IIf(destinationPath.EndsWith("\"), destinationPath, destinationPath & "\") & _filename, True)
            Next

            Return ""
        Catch ex As Exception
            Return "พบปัญหาในการคัดลอกไฟล์ : Application log :" & ex.ToString
        End Try

    End Function

    Public Sub CreateZIPFile(ByVal ExpName As String, ByVal Password As String, ByVal FileNamess As String())
        Try

            Dim fsOut As FileStream = File.Create(ExpName)
            Dim zipStream As New ZipOutputStream(fsOut)

            zipStream.SetLevel(3)
            'zipStream.Password = Password

            For Each filename As String In FileNamess
                Dim fi As New FileInfo(filename)
                Dim entryName As String = ZipEntry.CleanName(filename)
                Dim newEntry As New ZipEntry(entryName)
                newEntry.DateTime = fi.LastWriteTime
                newEntry.Size = fi.Length
                zipStream.PutNextEntry(newEntry)
                Dim buffer As Byte() = New Byte(4095) {}
                Using streamReader As FileStream = File.OpenRead(filename)
                    StreamUtils.Copy(streamReader, zipStream, buffer)
                End Using
                zipStream.CloseEntry()
            Next
            zipStream.IsStreamOwner = True
            zipStream.Close()
        Catch ex As Exception


        End Try
    End Sub
End Class
