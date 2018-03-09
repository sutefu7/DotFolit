Imports System.Globalization
Imports System.Windows.Controls.Primitives

Public Class TreeNodeKindsToStringConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert

        If TypeOf value IsNot TreeNodeKinds Then
            Return String.Empty
        End If

        Dim enumName = [Enum].GetName(GetType(TreeNodeKinds), value)
        Return enumName

    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function

End Class
