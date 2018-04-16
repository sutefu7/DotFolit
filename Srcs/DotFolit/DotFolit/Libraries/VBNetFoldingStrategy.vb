Imports ICSharpCode.AvalonEdit.Document
Imports ICSharpCode.AvalonEdit.Folding


Public Class VBNetFoldingStrategy

    Public Sub UpdateFoldings(manager As FoldingManager, document As TextDocument)

        Dim firstErrorOffset As Integer = -1
        Dim foldings = Me.CreateNewFoldings(document, firstErrorOffset)
        Dim sortedItems = foldings.OrderBy(Function(x) x.StartOffset)

        manager.UpdateFoldings(sortedItems, firstErrorOffset)

    End Sub

    ' Class, Method などコンテナ単位で折りたたむ開始位置、終了位置、折りたたんだ際の表示名を返却
    Public Iterator Function CreateNewFoldings(document As TextDocument, firstErrorOffset As Integer) As IEnumerable(Of NewFolding)

        Dim source = document.Text
        Dim walker = New FoldingSyntaxWalker
        walker.Parse(source)

        For Each item In walker.Items
            Yield New NewFolding With {.StartOffset = item.Item1, .EndOffset = item.Item2, .Name = item.Item3}
        Next

    End Function

End Class
