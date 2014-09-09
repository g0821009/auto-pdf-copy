'version 1.00 2014.9.9.1103 first Release

Imports System.IO

Module Module1

    Sub Main()
        'Dim sourceDirName As String = "C:\Users\t-uehara\Desktop\test\From"
        Dim sourceDirName As String = "\\192.168.160.105\部品加工作業要領書"
        Dim destDirName As String = "C:\Users\Public\Desktop\test\To\部品加工作業要領書"

        Dim timeCount As Double = Timer

        Console.WriteLine("部品加工作業要領書 差分ファイルコピー開始")

        fileCopy(sourceDirName, destDirName)
        Console.WriteLine("部品加工作業要領書コピー終了 実行時間(s):" & Timer() - timeCount)

        sourceDirName = "\\192.168.160.105\承認pdfデータ領域"
        destDirName = "C:\Users\Public\Desktop\test\To\承認pdfデータ領域"

        Console.WriteLine("承認pdfデータ領域 差分ファイルコピー開始")

        fileCopy(sourceDirName, destDirName)

        Console.WriteLine("finish")
        Console.WriteLine("承認pdfデータ領域コピー終了 実行時間(s):" & Timer() - timeCount)
        MsgBox("差分PDFファイルのコピーが終了しました")
    End Sub

    Sub fileCopy(ByVal sourceDirName As String, ByVal destDirName As String)
        'コピー先のディレクトリ名の末尾に"\"をつける
        If destDirName(destDirName.Length - 1) <> Path.DirectorySeparatorChar Then
            destDirName = destDirName + Path.DirectorySeparatorChar
        End If

        'コピー元のディレクトリにあるファイルをコピー
        Dim files As String() = Directory.GetFiles(sourceDirName, "*.pdf", SearchOption.TopDirectoryOnly)
        Console.WriteLine("get file name finished")
        Dim f As String
        Dim i As Long = 0
        Dim maxDim As Long = UBound(files)
        Dim tmpTime As Double = Timer

        Console.CursorVisible = False
        For Each f In files
            Dim destFileName As String = destDirName + Path.GetFileName(f)
            'コピー先にファイルが存在しない、
            '存在してもコピー元より更新日時が古い時はコピーする
            If Not File.Exists(destFileName) OrElse File.GetLastWriteTime(destFileName) < File.GetLastWriteTime(f) Then
                File.Copy(f, destFileName, True)
                Console.WriteLine("    copy:" & Path.GetFileName(f))
            End If
            i = i + 1
            If Timer - tmpTime > 2 Then
                Console.Write(String.Format("{0, 4:p} ", i / maxDim))
                Console.WriteLine(i & "/" & maxDim)
                ' カーソル位置を初期化
                Console.SetCursorPosition(0, Console.CursorTop - 1)
                tmpTime = Timer
            End If
        Next
        Console.CursorVisible = True
    End Sub

End Module
