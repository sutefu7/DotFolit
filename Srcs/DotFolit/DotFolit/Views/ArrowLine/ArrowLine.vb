' ------------------------------------------
'  ArrowLine.cs (c) 2007 by Charles Petzold
' ------------------------------------------

' このソースは、C# ソースを VB ソースに、自前で変換したものです。
' 上記オリジナルのライセンスに属します。
' http://www.charlespetzold.com/blog/2007/04/191200.html

' 同じチームの人？
' https://denisvuyka.wordpress.com/2007/10/28/wpf-diagramming-lines-pointing-at-the-center-of-the-element-calculating-angles-for-render-transforms/


Imports System
Imports System.Windows
Imports System.Windows.Media


Namespace Petzold.Media2D

    ''' <summary>
    ''' Draws a straight line between two points with 
    ''' optional arrows on the ends.
    ''' </summary>
    Public Class ArrowLine
        Inherits ArrowLineBase

        ''' <summary>
        ''' Identifies the X1 dependency property.
        ''' </summary>
        Public Shared ReadOnly X1Property As DependencyProperty =
            DependencyProperty.Register(
            "X1",
            GetType(Double),
            GetType(ArrowLine),
            New FrameworkPropertyMetadata(0.0, FrameworkPropertyMetadataOptions.AffectsMeasure))

        ''' <summary>
        ''' Gets or sets the x-coordinate of the ArrowLine start point.
        ''' </summary>
        ''' <returns></returns>
        Public Property X1 As Double
            Get
                Return CType(Me.GetValue(X1Property), Double)
            End Get
            Set(value As Double)
                Me.SetValue(X1Property, value)
            End Set
        End Property

        ''' <summary>
        ''' Identifies the Y1 dependency property.
        ''' </summary>
        Public Shared ReadOnly Y1Property As DependencyProperty =
            DependencyProperty.Register(
            "Y1",
            GetType(Double),
            GetType(ArrowLine),
            New FrameworkPropertyMetadata(0.0, FrameworkPropertyMetadataOptions.AffectsMeasure))

        ''' <summary>
        ''' Gets or sets the y-coordinate of the ArrowLine start point.
        ''' </summary>
        ''' <returns></returns>
        Public Property Y1 As Double
            Get
                Return CType(Me.GetValue(Y1Property), Double)
            End Get
            Set(value As Double)
                Me.SetValue(Y1Property, value)
            End Set
        End Property

        ''' <summary>
        ''' Identifies the X2 dependency property.
        ''' </summary>
        Public Shared ReadOnly X2Property As DependencyProperty =
            DependencyProperty.Register(
            "X2",
            GetType(Double),
            GetType(ArrowLine),
            New FrameworkPropertyMetadata(0.0, FrameworkPropertyMetadataOptions.AffectsMeasure))

        ''' <summary>
        ''' Gets or sets the x-coordinate of the ArrowLine end point.
        ''' </summary>
        ''' <returns></returns>
        Public Property X2 As Double
            Get
                Return CType(Me.GetValue(X2Property), Double)
            End Get
            Set(value As Double)
                Me.SetValue(X2Property, value)
            End Set
        End Property

        ''' <summary>
        ''' Identifies the Y2 dependency property.
        ''' </summary>
        Public Shared ReadOnly Y2Property As DependencyProperty =
            DependencyProperty.Register(
            "Y2",
            GetType(Double),
            GetType(ArrowLine),
            New FrameworkPropertyMetadata(0.0, FrameworkPropertyMetadataOptions.AffectsMeasure))

        ''' <summary>
        ''' Gets or sets the y-coordinate of the ArrowLine end point.
        ''' </summary>
        ''' <returns></returns>
        Public Property Y2 As Double
            Get
                Return CType(Me.GetValue(Y2Property), Double)
            End Get
            Set(value As Double)
                Me.SetValue(Y2Property, value)
            End Set
        End Property

        ''' <summary>
        ''' Gets a value that represents the Geometry of the ArrowLine.
        ''' </summary>
        ''' <returns></returns>
        Protected Overrides ReadOnly Property DefiningGeometry As Geometry
            Get

                ' Clear out the PathGeometry.
                pathgeo.Figures.Clear()

                ' Define a single PathFigure with the points.
                pathfigLine.StartPoint = New Point(X1, Y1)
                polysegLine.Points.Clear()
                polysegLine.Points.Add(New Point(X2, Y2))
                pathgeo.Figures.Add(pathfigLine)

                ' Call the base property to add arrows on the ends.
                Return MyBase.DefiningGeometry

            End Get
        End Property

    End Class

End Namespace
