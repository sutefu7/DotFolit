Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Public Class ContainerIfToGridConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert

        Dim items = TryCast(value, List(Of BaseTemplateModel))
        If (items Is Nothing) OrElse (items.Count = 0) Then
            Return Nothing
        End If

        ' 行数は１つ、列数は条件分岐の数分を定義
        Dim grid1 = New Grid
        grid1.RowDefinitions.Add(New RowDefinition With {.Height = New GridLength(1, GridUnitType.Star)})
        items.ForEach(Sub(x) grid1.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)}))

        ' 各セルに、土台となる ContentControl と条件分岐１つ分のデータをセット
        For i As Integer = 0 To items.Count - 1

            Dim contentcontrol1 = New ContentControl
            contentcontrol1.Content = items(i)
            contentcontrol1.SetValue(Grid.RowProperty, 0)
            contentcontrol1.SetValue(Grid.ColumnProperty, i)
            grid1.Children.Add(contentcontrol1)

        Next

        Return grid1

    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function

End Class
