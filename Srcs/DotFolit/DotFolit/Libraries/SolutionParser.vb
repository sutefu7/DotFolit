Imports System.IO


Public Class SolutionParser

    Public Iterator Function GetProjectFiles(solutionFile As String) As IEnumerable(Of String)

        Dim slnDir = Path.GetDirectoryName(solutionFile)
        Dim lines = File.ReadAllLines(solutionFile, EncodeResolver.GetEncoding(solutionFile))

        For Each line In lines

            If line.StartsWith("Project(") Then

                ' Project("{F184B08F-C81C-45F6-A57F-5ABD9991F28F}") = "ConsoleApplication28", "ConsoleApplication28\ConsoleApplication28.vbproj", "{A0D7A93C-B77F-4C8D-BD9D-7389C0495D33}"
                Dim x = line.Substring(line.IndexOf("=") + 1)
                x = x.Trim()

                ' ２つ目のデータとなるプロジェクトのパスを取得、囲っているダブルコーテーションは除去
                Dim xs = x.Split(","c)
                x = xs(1).Trim()
                x = x.Replace(ControlChars.Quote, String.Empty)

                Dim projectFile = Path.Combine(slnDir, x)
                If File.Exists(projectFile) Then
                    Yield projectFile
                End If

            End If

        Next

    End Function

    Public Iterator Function GetProjectDisplayNameAndFiles(slnFile As String) As IEnumerable(Of Tuple(Of String, String))

        Dim slnDir = Path.GetDirectoryName(slnFile)
        Dim lines = File.ReadAllLines(slnFile, EncodeResolver.GetEncoding(slnFile))

        For Each line In lines

            If line.StartsWith("Project(") Then

                ' Project("{F184B08F-C81C-45F6-A57F-5ABD9991F28F}") = "ConsoleApplication28", "ConsoleApplication28\ConsoleApplication28.vbproj", "{A0D7A93C-B77F-4C8D-BD9D-7389C0495D33}"
                Dim x = line.Substring(line.IndexOf("=") + 1)
                x = x.Trim()

                ' １つ目の表示名を取得
                Dim xs = x.Split(","c)
                x = xs(0).Trim()
                x = x.Replace(ControlChars.Quote, String.Empty)
                Dim projectName = x

                ' ２つ目のデータとなるプロジェクトのパスを取得、囲っているダブルコーテーションは除去
                x = xs(1).Trim()
                x = x.Replace(ControlChars.Quote, String.Empty)

                Dim projectFile = Path.Combine(slnDir, x)
                If File.Exists(projectFile) Then
                    Yield Tuple.Create(projectName, projectFile)
                End If

            End If

        Next

    End Function

End Class
