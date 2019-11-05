Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Media

''' <summary>
''' Contains methods that are used to extract text out of a FrameworkElement object.
''' </summary>
Module GemBoxDocumentHelper
    <Runtime.CompilerServices.Extension>
    Function ToText(ByVal root As FrameworkElement) As String
        Dim builder As New StringBuilder()

        For Each visual In root.GetSelfAndDescendants().OfType(Of DrawingVisual)()
            Dim previousRun As GlyphRun = Nothing

            ' Order runs first vertically (Y), then horizontally (X).
            For Each currentRun In visual.Drawing _
                .GetSelfAndDescendants() _
                .OfType(Of GlyphRunDrawing)() _
                .Select(Function(glyph) glyph.GlyphRun) _
                .OrderBy(Function(run) run.BaselineOrigin.Y) _
                .ThenBy(Function(run) run.BaselineOrigin.X)

                If previousRun IsNot Nothing Then
                    ' If base-line of current text segment is left from base-line of previous text segment, then assume that it is new line.
                    If currentRun.BaselineOrigin.X <= previousRun.BaselineOrigin.X Then
                        builder.AppendLine()
                    Else
                        Dim currentRect As Rect = currentRun.ComputeAlignmentBox()
                        Dim previousRect As Rect = previousRun.ComputeAlignmentBox()

                        Dim spaceWidth As Double = currentRun.BaselineOrigin.X + currentRect.Left - previousRun.BaselineOrigin.X - previousRect.Right
                        Dim spaceHeight As Double = (currentRect.Height + previousRect.Height) / 2

                        ' If space between successive text segments has width greater than a sixth of its height, then assume that it is a word (add a space).
                        If spaceWidth > spaceHeight / 6 Then builder.Append(" "c)
                    End If
                End If

                builder.Append(currentRun.Characters.ToArray())
                previousRun = currentRun
            Next
        Next

        Return builder.ToString()
    End Function

    <Runtime.CompilerServices.Extension>
    Private Iterator Function GetSelfAndDescendants(ByVal parent As DependencyObject) As IEnumerable(Of DependencyObject)
        Yield parent

        Dim count = VisualTreeHelper.GetChildrenCount(parent)
        For i = 0 To count - 1
            For Each descendant In VisualTreeHelper.GetChild(parent, i).GetSelfAndDescendants()
                Yield descendant
            Next
        Next
    End Function

    <Runtime.CompilerServices.Extension>
    Private Iterator Function GetSelfAndDescendants(ByVal parent As DrawingGroup) As IEnumerable(Of Drawing)
        Yield parent

        For Each child In parent.Children
            Dim group = TryCast(child, DrawingGroup)
            If group IsNot Nothing Then
                For Each descendant In group.GetSelfAndDescendants()
                    Yield descendant
                Next
            Else
                Yield child
            End If
        Next
    End Function
End Module