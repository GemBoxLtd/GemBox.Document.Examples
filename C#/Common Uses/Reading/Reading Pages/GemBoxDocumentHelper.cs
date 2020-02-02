using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;

/// <summary>
/// Contains methods that are used to extract text out of a FrameworkElement object.
/// </summary>
public static class GemBoxDocumentHelper
{
    public static string ToText(this FrameworkElement root)
    {
        var builder = new StringBuilder();

        foreach (var visual in root.GetSelfAndDescendants().OfType<DrawingVisual>())
        {
            GlyphRun previousRun = null;

            // Order runs first vertically (Y), then horizontally (X).
            foreach (var currentRun in visual.Drawing
                .GetSelfAndDescendants()
                .OfType<GlyphRunDrawing>()
                .Select(glyph => glyph.GlyphRun)
                .OrderBy(run => run.BaselineOrigin.Y)
                .ThenBy(run => run.BaselineOrigin.X))
            {
                if (previousRun != null)
                {
                    // If base-line of current text segment is left from base-line of previous text segment, then assume that it is new line.
                    if (currentRun.BaselineOrigin.X <= previousRun.BaselineOrigin.X)
                    {
                        builder.AppendLine();
                    }
                    else
                    {
                        Rect currentRect = currentRun.ComputeAlignmentBox();
                        Rect previousRect = previousRun.ComputeAlignmentBox();

                        double spaceWidth = currentRun.BaselineOrigin.X + currentRect.Left - previousRun.BaselineOrigin.X - previousRect.Right;
                        double spaceHeight = (currentRect.Height + previousRect.Height) / 2;

                        // If space between successive text segments has width greater than a sixth of its height, then assume that it is a word (add a space).
                        if (spaceWidth > spaceHeight / 6)
                            builder.Append(' ');
                    }
                }

                builder.Append(currentRun.Characters.ToArray());
                previousRun = currentRun;
            }
        }

        return builder.ToString();
    }

    private static IEnumerable<DependencyObject> GetSelfAndDescendants(this DependencyObject parent)
    {
        yield return parent;

        for (int i = 0, count = VisualTreeHelper.GetChildrenCount(parent); i < count; i++)
            foreach (var descendant in VisualTreeHelper.GetChild(parent, i).GetSelfAndDescendants())
                yield return descendant;
    }

    private static IEnumerable<Drawing> GetSelfAndDescendants(this DrawingGroup parent)
    {
        yield return parent;

        foreach (var child in parent.Children)
        {
            var group = child as DrawingGroup;
            if (group != null)
                foreach (var descendant in group.GetSelfAndDescendants())
                    yield return descendant;
            else
                yield return child;
        }
    }
}