using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Logging;

namespace Lefty.Powerpoint;

/// <summary />
public class PowerpointCleanerService
{
    private readonly ILogger<PowerpointCleanerService> _logger;


    /// <summary />
    public PowerpointCleanerService( ILogger<PowerpointCleanerService> logger )
    {
        _logger = logger;
    }


    /// <summary />
    public async Task<PowerpointCleanResult> CleanAsync( string filePath, string outputPath )
    {
        await Task.Yield();


        /*
         * Create copy
         */
        File.Copy( filePath, outputPath, true );


        /*
         * Manipulate
         */
        var res = Clean( outputPath );

        return res;
    }


    /// <summary />
    private PowerpointCleanResult Clean( string filePath )
    {
        var res = new PowerpointCleanResult();


        using var pptx = PresentationDocument.Open( filePath, true );

        if ( pptx.PresentationPart == null )
            return res;


        //
        var presentation = pptx.PresentationPart.Presentation;
        var slideIdList = presentation.SlideIdList;

        if ( slideIdList == null )
            return res;


        // Check number of slide masters
        res.SlideMasters = pptx.PresentationPart.SlideMasterParts.Count();


        // Get all slide IDs to process (iterate backwards to safely remove)
        var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();

        for ( int i = slideIds.Count - 1; i >= 0; i-- )
        {
            SlideId slideId = slideIds[ i ];
            var slidePart = (SlidePart) pptx.PresentationPart.GetPartById( slideId.RelationshipId! );


            // Get slide title
            var slideTitle = GetSlideTitle( slidePart );


            // Check if slide is hidden
            bool isHidden = slidePart.Slide.Show != null && slidePart.Slide.Show.HasValue && !slidePart.Slide.Show.Value;

            if ( isHidden == true )
            {
                _logger.LogInformation( "#{Number}: Hidden, removing slide {Title}...", i, slideTitle );

                // Remove hidden slide
                slideId.Remove();
                pptx.PresentationPart.DeletePart( slidePart );

                res.RemovedSlides++;
                continue;
            }


            // Remove comments from visible slides
            if ( slidePart.SlideCommentsPart != null )
            {
                _logger.LogInformation( "#{Number}: Has comments, removing from {Title}...", i, slideTitle );

                slidePart.DeletePart( slidePart.SlideCommentsPart );
            }


            // Remove notes from visible slides
            if ( slidePart.NotesSlidePart != null )
            {
                _logger.LogInformation( "#{Number}: Has notes, clearing from {Title}...", i, slideTitle );

                slidePart.DeletePart( slidePart.NotesSlidePart );
            }
        }


        // Also remove comment authors part if it exists
        if ( pptx.PresentationPart.CommentAuthorsPart != null )
        {
            _logger.LogInformation( "Removing comment authors..." );

            pptx.PresentationPart.DeletePart( pptx.PresentationPart.CommentAuthorsPart );
        }

        presentation.Save();

        return res;
    }


    /// <summary />
    private static string GetSlideTitle( SlidePart slidePart )
    {
        // Look for the title shape (typically the first shape with a title placeholder type)
        var slide = slidePart.Slide;
        var shapes = slide.CommonSlideData?.ShapeTree?.Elements<Shape>();

        if ( shapes == null )
            return "[No Title]";

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
            if ( !string.IsNullOrWhiteSpace( text ) )
                return text;
        }

        return "[No Title]";
    }


    /// <summary />
    private static string GetTextFromShape( Shape shape )
    {
        var textBody = shape.TextBody;
        if ( textBody == null )
            return string.Empty;

        var paragraphs = textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>();
        var textParts = new List<string>();

        foreach ( var paragraph in paragraphs )
        {
            var runs = paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>();
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