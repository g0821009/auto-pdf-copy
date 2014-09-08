Imports System.IO

Module Module1

    Sub Main()
        'Dim sourceDirName As String = "C:\Users\t-uehara\Desktop\test\From"
        Dim sourceDirName As String = "\\192.168.160.105\部品加工作業要領書"
        Dim destDirName As String = "C:\Users\t-uehara\Desktop\test\To\部品加工作業要領書"

        fileCopy(sourceDirName, destDirName)

        sourceDirName = "\\192.168.160.105\承認pdfデータ領域"
        destDirName = "C:\Users\t-uehara\Desktop\test\To\承認pdfデータ領域"

        fileCopy(sourceDirName, destDirName)


        Debug.WriteLine("finish")
        MsgBox("finish")
    End Sub

    Sub fileCopy(ByVal sourceDirName As String, ByVal destDirName As String)
        'コピー先のディレクトリ名の末尾に"\"をつける
        If destDirName(destDirName.Length - 1) <> Path.DirectorySeparatorChar Then
            destDirName = destDirName + Path.DirectorySeparatorChar
        End If

        'コピー元のディレクトリにあるファイルをコピー
        Dim files As String() = Directory.GetFiles(sourceDirName, "*.pdf", SearchOption.TopDirectoryOnly)
        Debug.WriteLine("get file name finished")
        Dim f As String
        For Each f In files
            Dim destFileName As String = destDirName + Path.GetFileName(f)
            'コピー先にファイルが存在し、
            'コピー元より更新日時が古い時はコピーする
            If Not File.Exists(destFileName) OrElse File.GetLastWriteTime(destFileName) < File.GetLastWriteTime(f) Then
                'File.Copy(f, destFileName, True)
                Debug.WriteLine("  copy " & destFileName)
            End If
        Next
    End Sub

End Module
