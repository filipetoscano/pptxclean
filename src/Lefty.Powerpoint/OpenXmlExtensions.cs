using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;

namespace Lefty.Powerpoint;

/// <summary />
internal static class OpenXmlExtensions
{
    /// <summary>
    /// Gets whether a slide is hidden or not.
    /// </summary>
    /// <param name="slide">Slide.</param>
    /// <returns>True if slide is hidden, false otherwise.</returns>
    internal static bool IsHidden( this SlidePart slide )
    {
        if ( slide.Slide.Show == null )
            return false;

        if ( slide.Slide.Show.HasValue == false )
            return false;

        if ( slide.Slide.Show.Value == false )
            return true;

        return false;
    }


    /// <summary>
    /// Gets the slide title.
    /// </summary>
    /// <param name="slide">Slide.</param>
    /// <returns>Title slide, or null if not found.</returns>
    internal static string? SlideTitle( this SlidePart slide )
    {
        // Look for the title shape (typically the first shape with a title placeholder type)
        var s = slide.Slide;
        var shapes = s.CommonSlideData?.ShapeTree?.Elements<Shape>();

        if ( shapes == null )
            return null;

        foreach ( var shape in shapes )
        {
            // Check if this shape is a title placeholder
            var placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;

            if ( placeholderShape != null )
            {
                // Type "title" or "ctrTitle" indicates a title placeholder
                var placeholderType = placeholderShape.Type?.Value;

                if ( placeholderType == PlaceholderValues.Title ||
                     placeholderType == PlaceholderValues.CenteredTitle )
                {
                    // Extract text from the shape
                    var text = GetTextFromShape( shape );

                    if ( !string.IsNullOrWhiteSpace( text ) )
                        return text;
                }
            }
        }

        // If no title placeholder found, try to get the first text shape
        foreach ( var shape in shapes )
        {
            var text = GetTextFromShape( shape );

            if ( string.IsNullOrWhiteSpace( text ) == false )
                return text;
        }

        return null;
    }


    /// <summary />
    internal static string GetTextFromShape( this Shape shape )
    {
        var textBody = shape.TextBody;

        if ( textBody == null )
            return string.Empty;

        var paragraphs = textBody.Elements<Paragraph>();
        var textParts = new List<string>();

        foreach ( var paragraph in paragraphs )
        {
            var runs = paragraph.Elements<Run>();
            foreach ( var run in runs )
            {
                var text = run.Text?.Text;
                if ( !string.IsNullOrEmpty( text ) )
                    textParts.Add( text );
            }
        }

        return string.Join( "", textParts ).Trim();
    }
}