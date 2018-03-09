' ------------------------------------------
'  ArrowEnds.cs (c) 2007 by Charles Petzold
' ------------------------------------------

' このソースは、C# ソースを VB ソースに、自前で変換したものです。
' 上記オリジナルのライセンスに属します。
' http://www.charlespetzold.com/blog/2007/04/191200.html

' 同じチームの人？
' https://denisvuyka.wordpress.com/2007/10/28/wpf-diagramming-lines-pointing-at-the-center-of-the-element-calculating-angles-for-render-transforms/


Imports System


Namespace Petzold.Media2D

    ''' <summary>
    ''' Indicates which end of the line has an arrow.
    ''' </summary>
    <Flags()>
    Public Enum ArrowEnds
        None = 0
        Start = 1
        [End] = 2
        Both = 3
    End Enum

End Namespace