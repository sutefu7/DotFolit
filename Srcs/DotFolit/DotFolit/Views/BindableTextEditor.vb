Imports ICSharpCode.AvalonEdit
Imports ICSharpCode.AvalonEdit.Editing
Imports ICSharpCode.AvalonEdit.Folding
Imports ICSharpCode.AvalonEdit.Document
Imports ICSharpCode.AvalonEdit.Highlighting


Public Class BindableTextEditor
    Inherits TextEditor

#Region "SourceCode 依存関係プロパティ"

    Public Shared ReadOnly SourceCodeProperty As DependencyProperty =
        DependencyProperty.Register(
        "SourceCode",
        GetType(String),
        GetType(BindableTextEditor),
        New PropertyMetadata(String.Empty, AddressOf OnSourceCodePropertyChanged))

    Private Shared Sub OnSourceCodePropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim ctrl = TryCast(d, BindableTextEditor)
        ctrl.SourceCode = CStr(e.NewValue)

    End Sub

    ' ※ TextEditor.Document.Text の方を対象にしないでください。
    ' TextEditor.Text 内で初期化処理をおこなっているため、こちらを対象にしています。
    ' 詳しくは、https://github.com/icsharpcode/AvalonEdit/blob/master/ICSharpCode.AvalonEdit/TextEditor.cs
    Public Property SourceCode As String
        Get
            Return MyBase.Text
        End Get
        Set(value As String)

            MyBase.Text = value

            ' ソースコードがセットされたので、VB.NET 言語用の折りたたみルールを適用
            ' ※（雑談）本ツールでは（仕様上問題無いので）何もしていないが、アプリの作りによっては、いったん FoldingManager.Uninstall した後で、Install し直す手順が必要になる場合があるみたい
            Dim strategy = New VBNetFoldingStrategy
            Dim manager = FoldingManager.Install(MyBase.TextArea)
            strategy.UpdateFoldings(manager, MyBase.Document)

        End Set
    End Property

#End Region

#Region "SourceFile 依存関係プロパティ"

    Public Shared ReadOnly SourceFileProperty As DependencyProperty =
        DependencyProperty.Register(
        "SourceFile",
        GetType(String),
        GetType(BindableTextEditor),
        New PropertyMetadata(String.Empty, AddressOf OnSourceFilePropertyChanged))

    Private Shared Sub OnSourceFilePropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim ctrl = TryCast(d, BindableTextEditor)
        ctrl.SourceFile = CStr(e.NewValue)

    End Sub

    Public Property SourceFile As String
        Get
            Return MyBase.Document.FileName
        End Get
        Set(value As String)
            MyBase.Document.FileName = value
        End Set
    End Property

#End Region

#Region "CaretLocation 依存関係プロパティ"

    Public Shared ReadOnly CaretLocationProperty As DependencyProperty =
        DependencyProperty.Register(
        "CaretLocation",
        GetType(TextLocation),
        GetType(BindableTextEditor),
        New PropertyMetadata(AddressOf OnCaretLocationPropertyChanged))

    Private Shared Sub OnCaretLocationPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim ctrl = TryCast(d, BindableTextEditor)
        ctrl.CaretLocation = CType(e.NewValue, TextLocation)

    End Sub

    Public Property CaretLocation As TextLocation
        Get
            Return MyBase.TextArea.Caret.Location
        End Get
        Set(value As TextLocation)

            ' Caret_PositionChanged イベントが、２回連続で発生してしまう不具合の対応
            If MyBase.TextArea.Caret.Location = value Then
                Return
            End If

            MyBase.TextArea.Caret.Location = value

        End Set
    End Property

#End Region

#Region "CaretOffset 依存関係プロパティ"

    Public Shared ReadOnly CaretOffsetProperty As DependencyProperty =
        DependencyProperty.Register(
        "CaretOffset",
        GetType(Integer),
        GetType(BindableTextEditor),
        New PropertyMetadata(AddressOf OnCaretOffsetPropertyChanged))

    Private Shared Sub OnCaretOffsetPropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)

        Dim ctrl = TryCast(d, BindableTextEditor)
        ctrl.CaretOffset = CInt(e.NewValue)

        ' キャレット位置（メンバー定義位置）までスクロールが見えるようにスクロール
        Dim jumpLine = ctrl.Document.GetLineByOffset(ctrl.CaretOffset).LineNumber
        ctrl.ScrollToLine(jumpLine)

    End Sub

    ' 今のところ、TextEditor.CaretOffset プロパティは、TextEditor.TextArea.Caret.Offset をラップしただけのプロパティになっているが、
    ' Text プロパティみたいに、将来の機能修正を考慮して TextEditor.CaretOffset の方をラップする
    ' また、継承元クラスに同名のプロパティがあるため、継承元クラス側のプロパティは隠した
    ' 依存関係プロパティと対になる CLR プロパティを定義しなくてはいけない仕様みたいなので定義している（これを書かないと xaml 上で見えない？）が、しなくてもいいのであれば、消したい（他の依存関係プロパティも同じ）
    Public Shadows Property CaretOffset As Integer
        Get
            Return MyBase.CaretOffset
        End Get
        Set(value As Integer)

            ' Caret_PositionChanged イベントが、２回連続で発生してしまう不具合の対応　
            If MyBase.CaretOffset = value Then
                Return
            End If

            MyBase.CaretOffset = value

        End Set
    End Property

#End Region

#Region "CaretPositionChanged ルーティングイベント"

    Public Shared ReadOnly CaretPositionChangedEvent As RoutedEvent =
        EventManager.RegisterRoutedEvent(
        "CaretPositionChanged",
        RoutingStrategy.Bubble,
        GetType(RoutedEventHandler),
        GetType(BindableTextEditor))

    Public Custom Event CaretPositionChanged As RoutedEventHandler
        AddHandler(value As RoutedEventHandler)
            Me.AddHandler(CaretPositionChangedEvent, value)
        End AddHandler
        RemoveHandler(value As RoutedEventHandler)
            Me.RemoveHandler(CaretPositionChangedEvent, value)
        End RemoveHandler
        RaiseEvent(sender As Object, e As RoutedEventArgs)
            Me.RaiseEvent(e)
        End RaiseEvent
    End Event

    Private Sub Caret_PositionChanged(sender As Object, e As EventArgs)

        ' VM 側のデータが更新されない不具合の対応
        ' CLR プロパティ -> 依存関係プロパティへの同期がおこなわれていないため、更新後の値を同期する
        Me.SetValue(CaretLocationProperty, Me.CaretLocation)
        Me.SetValue(CaretOffsetProperty, Me.CaretOffset)

        Dim newEventArgs = New RoutedEventArgs(CaretPositionChangedEvent, Me)
        Me.RaiseEvent(newEventArgs)

    End Sub

#End Region

#Region "コンストラクタ"

    Public Sub New()
        MyBase.New()

        ' 各種設定
        Me.IsReadOnly = True
        Me.ShowLineNumbers = True
        Me.SyntaxHighlighting = HighlightingManager.Instance.GetDefinition("VB")
        Me.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        Me.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto

        Me.Options.ConvertTabsToSpaces = True
        Me.Options.HighlightCurrentLine = True

        ' キャレット移動イベントを購読
        AddHandler Me.TextArea.Caret.PositionChanged, AddressOf Me.Caret_PositionChanged

    End Sub

#End Region

End Class
