' ------------------------------------------
'  ArrowLineBase.cs (c) 2007 by Charles Petzold
' ------------------------------------------

' このソースは、C# ソースを VB ソースに、自前で変換したものです。
' 上記オリジナルのライセンスに属します。
' http://www.charlespetzold.com/blog/2007/04/191200.html

' 同じチームの人？
' https://denisvuyka.wordpress.com/2007/10/28/wpf-diagramming-lines-pointing-at-the-center-of-the-element-calculating-angles-for-render-transforms/


Imports System
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Shapes


Namespace Petzold.Media2D

    ''' <summary>
    ''' Provides a base class for ArrowLine and ArrowPolyline.
    ''' This class is abstract.
    ''' </summary>
    Public MustInherit Class ArrowLineBase
        Inherits Shape

        Protected pathgeo As PathGeometry
        Protected pathfigLine As PathFigure
        Protected polysegLine As PolyLineSegment

        Private pathfigHead1 As PathFigure
        Private polysegHead1 As PolyLineSegment
        Private pathfigHead2 As PathFigure
        Private polysegHead2 As PolyLineSegment

        ''' <summary>
        ''' Identifies the ArrowAngle dependency property.
        ''' </summary>
        Public Shared ReadOnly ArrowAngleProperty As DependencyProperty =
            DependencyProperty.Register(
            "ArrowAngle",
            GetType(Double),
            GetType(ArrowLineBase),
            New FrameworkPropertyMetadata(45.0, FrameworkPropertyMetadataOptions.AffectsMeasure))

        ''' <summary>
        ''' Gets or sets the angle between the two sides of the arrowhead.
        ''' </summary>
        ''' <returns></returns>
        Public Property ArrowAngle As Double
            Get
                Return CType(Me.GetValue(ArrowAngleProperty), Double)
            End Get
            Set(value As Double)
                Me.SetValue(ArrowAngleProperty, value)
            End Set
        End Property

        ''' <summary>
        ''' Identifies the ArrowLength dependency property.
        ''' </summary>
        Public Shared ReadOnly ArrowLengthProperty As DependencyProperty =
            DependencyProperty.Register(
            "ArrowLength",
            GetType(Double),
            GetType(ArrowLineBase),
            New FrameworkPropertyMetadata(12.0, FrameworkPropertyMetadataOptions.AffectsMeasure))

        ''' <summary>
        ''' Gets or sets the length of the two sides of the arrowhead.
        ''' </summary>
        ''' <returns></returns>
        Public Property ArrowLength As Double
            Get
                Return CType(Me.GetValue(ArrowLengthProperty), Double)
            End Get
            Set(value As Double)
                Me.SetValue(ArrowLengthProperty, value)
            End Set
        End Property

        ''' <summary>
        ''' Identifies the ArrowEnds dependency property.
        ''' </summary>
        ''' <remarks>オリジナルは ArrowEnds という名称ですが、列挙体型の名称と同名のため、変更しました。</remarks>
        Public Shared ReadOnly ArrowEndProperty As DependencyProperty =
            DependencyProperty.Register(
            "ArrowEnd",
            GetType(ArrowEnds),
            GetType(ArrowLineBase),
            New FrameworkPropertyMetadata(ArrowEnds.End, FrameworkPropertyMetadataOptions.AffectsMeasure))

        ''' <summary>
        ''' Gets or sets the property that determines which ends of the
        ''' line have arrows.
        ''' </summary>
        ''' <remarks>オリジナルは ArrowEnds という名称ですが、列挙体型の名称と同名のため、変更しました。</remarks>
        ''' <returns></returns>
        Public Property ArrowEnd As ArrowEnds
            Get
                Return CType(Me.GetValue(ArrowEndProperty), ArrowEnds)
            End Get
            Set(value As ArrowEnds)
                Me.SetValue(ArrowEndProperty, value)
            End Set
        End Property

        ''' <summary>
        ''' Identifies the IsArrowClosed dependency property.
        ''' </summary>
        Public Shared ReadOnly IsArrowClosedProperty As DependencyProperty =
            DependencyProperty.Register(
            "IsArrowClosed",
            GetType(Boolean),
            GetType(ArrowLineBase),
            New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.AffectsMeasure))

        ''' <summary>
        ''' Gets or sets the property that determines if the arrow head
        ''' is closed to resemble a triangle.
        ''' </summary>
        ''' <returns></returns>
        Public Property IsArrowClosed As Boolean
            Get
                Return CType(Me.GetValue(IsArrowClosedProperty), Boolean)
            End Get
            Set(value As Boolean)
                Me.SetValue(IsArrowClosedProperty, value)
            End Set
        End Property

        ''' <summary>
        ''' Initializes a new instance of ArrowLineBase.
        ''' </summary>
        Public Sub New()

            pathgeo = New PathGeometry

            pathfigLine = New PathFigure
            polysegLine = New PolyLineSegment
            pathfigLine.Segments.Add(polysegLine)

            pathfigHead1 = New PathFigure
            polysegHead1 = New PolyLineSegment
            pathfigHead1.Segments.Add(polysegHead1)

            pathfigHead2 = New PathFigure
            polysegHead2 = New PolyLineSegment
            pathfigHead2.Segments.Add(polysegHead2)

        End Sub

        ''' <summary>
        ''' Gets a value that represents the Geometry of the ArrowLine.
        ''' </summary>
        ''' <returns></returns>
        Protected Overrides ReadOnly Property DefiningGeometry As Geometry
            Get

                Dim count = polysegLine.Points.Count

                If count > 0 Then

                    ' Draw the arrow at the start of the line.
                    If (ArrowEnd And ArrowEnds.Start) = ArrowEnds.Start Then

                        Dim pt1 As Point = pathfigLine.StartPoint
                        Dim pt2 As Point = polysegLine.Points(0)
                        pathgeo.Figures.Add(CalculateArrow(pathfigHead1, pt2, pt1))

                    End If

                    ' Draw the arrow at the end of the line.
                    If (ArrowEnd And ArrowEnds.End) = ArrowEnds.End Then

                        Dim pt1 As Point = If(count = 1, pathfigLine.StartPoint, polysegLine.Points(count - 2))
                        Dim pt2 As Point = polysegLine.Points(count - 1)
                        pathgeo.Figures.Add(CalculateArrow(pathfigHead2, pt1, pt2))

                    End If

                End If

                Return pathgeo

            End Get
        End Property

        Private Function CalculateArrow(pathfig As PathFigure, pt1 As Point, pt2 As Point) As PathFigure

            Dim matx As Matrix = New Matrix
            Dim vect As Vector = pt1 - pt2
            vect.Normalize()
            vect *= ArrowLength

            Dim polyseg As PolyLineSegment = TryCast(pathfig.Segments(0), PolyLineSegment)
            polyseg.Points.Clear()
            matx.Rotate(ArrowAngle / 2)
            pathfig.StartPoint = pt2 + vect * matx
            polyseg.Points.Add(pt2)

            matx.Rotate(-ArrowAngle)
            polyseg.Points.Add(pt2 + vect * matx)
            pathfig.IsClosed = IsArrowClosed

            Return pathfig

        End Function

    End Class

End Namespace
