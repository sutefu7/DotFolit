Imports Livet


' 【WPF】ViewModelがINotifyPropertyChangedを実装していないとメモリリークする件
' http://aridai.net/article/?p=15
' とのことらしいので、継承を追加


Public Class BaseTemplateModel
    Inherits ViewModel

    Public Property Header As String
    Public Property Signature As String
    Public Property Children As List(Of BaseTemplateModel)

    Public Sub New()
        MyBase.New()

        Header = String.Empty
        Signature = String.Empty
        Children = New List(Of BaseTemplateModel)

    End Sub

End Class
