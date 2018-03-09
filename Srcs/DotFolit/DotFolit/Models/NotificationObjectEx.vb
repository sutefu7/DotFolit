Imports System.Collections.ObjectModel
Imports Livet

Public Class NotificationObjectEx
    Inherits NotificationObject

    ' NotificationObjectEx を継承しているクラスの値比較を、イコール演算子でできるように作成
    ' プロパティ値が同じであれば、違う参照先でも同じデータと判断する
    Public Shared Operator =(instance1 As NotificationObjectEx, instance2 As NotificationObjectEx) As Boolean

        ' どちらかが Nothing の場合
        If (instance1 Is Nothing) OrElse (instance2 Is Nothing) Then
            Return False
        End If

        ' プロパティ数が違う（＝継承先が違うクラス同士で、インスタンス比較しようとした）場合
        Dim aProperties = instance1.GetType().GetProperties()
        Dim bProperties = instance2.GetType().GetProperties()

        If aProperties.Count() <> bProperties.Count() Then
            Return False
        End If

        ' 各プロパティ同士の値が違う場合
        For i As Integer = 0 To aProperties.Count() - 1

            Dim aValue = aProperties(i).GetValue(instance1)
            Dim bValue = bProperties(i).GetValue(instance2)

            If Not aValue.Equals(bValue) Then
                Return False
            End If

        Next

        Return True

    End Operator

    Public Shared Operator <>(instance1 As NotificationObjectEx, instance2 As NotificationObjectEx) As Boolean
        Return Not (instance1 = instance2)
    End Operator

End Class
