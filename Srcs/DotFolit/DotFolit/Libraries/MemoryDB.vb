Imports System.Data
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic


Public NotInheritable Class MemoryDB

#Region "シングルトン"

    Private Shared _Instance As MemoryDB = Nothing

    Public Shared ReadOnly Property Instance As MemoryDB
        Get

            If _Instance Is Nothing Then
                _Instance = New MemoryDB
            End If

            Return _Instance

        End Get
    End Property

    Private Sub New()

        Me.DB = Me.CreateDataSetLayout()

    End Sub

#End Region

#Region "フィールド、プロパティ"

    Public Property DB As DataSet = Nothing
    Public Property SyntaxTreeItems As List(Of SyntaxTree) = Nothing
    Public Property CompilationItem As VisualBasicCompilation = Nothing


#End Region

#Region "メソッド"

    Public Function CreateDataSetLayout() As DataSet

        Dim ds = New DataSet
        Dim table As DataTable = Nothing

        ' Class,  ConsoleApp1,        ConsoleApp1.Class1,         "",        "",        ""
        ' Field,  ConsoleApp1.Class1, ConsoleApp1.Class1._Width,  "Integer", "",        ""
        ' Method, ConsoleApp1.Class1, ConsoleApp1.Class1.GetData, "",        "Integer", "0()"
        ' Method, ConsoleApp1.Class1, ConsoleApp1.Class1.SetData, "",        "Void",    "1(Integer)"

        ' NamespaceResolution テーブル
        table = ds.Tables.Add("NamespaceResolution")
        table.Columns.Add("DefineKind", GetType(String))
        table.Columns.Add("ContainerFullName", GetType(String))
        table.Columns.Add("DefineFullName", GetType(String))
        table.Columns.Add("DefineType", GetType(String))
        table.Columns.Add("ReturnType", GetType(String))
        table.Columns.Add("MethodArguments", GetType(String))
        table.Columns.Add("IsPartial", GetType(Boolean))
        table.Columns.Add("IsShared", GetType(Boolean))
        table.Columns.Add("SourceFile", GetType(String))
        table.Columns.Add("StartLength", GetType(Integer))
        table.Columns.Add("EndLength", GetType(Integer))
        table.Columns.Add("StartLineNumber", GetType(Integer))
        table.Columns.Add("EndLineNumber", GetType(Integer))

        ' LanguageConversion テーブル
        table = ds.Tables.Add("LanguageConversion")
        table.Columns.Add("VBNet", GetType(String))
        table.Columns.Add("NETFramework", GetType(String))

        table.Rows.Add(New Object() {"Boolean", "System.Boolean"})
        table.Rows.Add(New Object() {"Byte", "System.Byte"})
        table.Rows.Add(New Object() {"Char", "System.Char"})
        table.Rows.Add(New Object() {"Short", "System.Int16"})
        table.Rows.Add(New Object() {"Integer", "System.Int32"})
        table.Rows.Add(New Object() {"Long", "System.Int64"})
        table.Rows.Add(New Object() {"SByte", "System.SByte"})
        table.Rows.Add(New Object() {"UShort", "System.UInt16"})
        table.Rows.Add(New Object() {"UInteger", "System.UInt32"})
        table.Rows.Add(New Object() {"ULong", "System.UInt64"})
        table.Rows.Add(New Object() {"Decimal", "System.Decimal"})
        table.Rows.Add(New Object() {"Single", "System.Single"})
        table.Rows.Add(New Object() {"Double", "System.Double"})
        table.Rows.Add(New Object() {"Date", "System.DateTime"})
        table.Rows.Add(New Object() {"DateTime", "System.DateTime"})
        table.Rows.Add(New Object() {"Object", "System.Object"})
        table.Rows.Add(New Object() {"String", "System.String"})
        table.Rows.Add(New Object() {"IntPtr", "System.IntPtr"})
        table.Rows.Add(New Object() {"UIntPtr", "System.UIntPtr"})

        table.Rows.Add(New Object() {"Short", "Int16"})
        table.Rows.Add(New Object() {"Integer", "Int32"})
        table.Rows.Add(New Object() {"Long", "Int64"})
        table.Rows.Add(New Object() {"UShort", "UInt16"})
        table.Rows.Add(New Object() {"UInteger", "UInt32"})
        table.Rows.Add(New Object() {"ULong", "UInt64"})

        table.Rows.Add(New Object() {"+", "op_Addition"})
        table.Rows.Add(New Object() {"-", "op_Subtraction"})
        table.Rows.Add(New Object() {"*", "op_Multiply"})
        table.Rows.Add(New Object() {"/", "op_Division"})
        table.Rows.Add(New Object() {"&", "op_Concatenate"})
        table.Rows.Add(New Object() {"<=", "op_LessThanOrEqual"})      ' LessThan にひっかかってしまい誤検知してしまうため、先に登録してしまう
        table.Rows.Add(New Object() {">=", "op_GreaterThanOrEqual"})   ' GreaterThan にひっかかってしまい誤検知してしまうため、先に登録してしまう
        table.Rows.Add(New Object() {"<", "op_LessThan"})
        table.Rows.Add(New Object() {">", "op_GreaterThan"})
        table.Rows.Add(New Object() {"<<", "op_LeftShift"})
        table.Rows.Add(New Object() {">>", "op_RightShift"})
        table.Rows.Add(New Object() {"<>", "op_Inequality"})
        table.Rows.Add(New Object() {"=", "op_Equality"})
        table.Rows.Add(New Object() {"\", "op_IntegerDivision"})
        table.Rows.Add(New Object() {"^", "op_Exponent"})
        table.Rows.Add(New Object() {"And", "op_BitwiseAnd"})
        table.Rows.Add(New Object() {"IsFalse", "op_False"})
        table.Rows.Add(New Object() {"IsTrue", "op_True"})
        table.Rows.Add(New Object() {"Like", "op_Like"})
        table.Rows.Add(New Object() {"Mod", "op_Modulus"})
        table.Rows.Add(New Object() {"Not", "op_OnesComplement"})
        table.Rows.Add(New Object() {"Or", "op_BitwiseOr"})
        table.Rows.Add(New Object() {"Xor", "op_ExclusiveOr"})
        table.Rows.Add(New Object() {"CType", "op_Explicit"})
        table.Rows.Add(New Object() {"CType", "op_Implicit"})

        Return ds

    End Function

#End Region

End Class
