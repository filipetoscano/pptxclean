using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Logging;
using System.Text.RegularExpressions;

namespace Lefty.Powerpoint;

/// <summary />
public class PowerpointService
{
    private readonly ILogger<PowerpointService> _logger;


    /// <summary />
    public PowerpointService( ILogger<PowerpointService> logger )
    {
        _logger = logger;
    }


    /// <summary />
    public async Task CheckAsync( string filePath )
    {
        await Task.Yield();

        throw new NotImplementedException();
    }


    /// <summary />
    public Task CleanAsync( string filePath )
    {
        Mutate( filePath, null );

        return Task.CompletedTask;
    }


    /// <summary />
    public async Task<PresentationSummary> GetAsync( string filePath )
    {
        await Task.Yield();

        var ps = new PresentationSummary();


        /*
         * 
         */
        using var pptx = PresentationDocument.Open( filePath, true );

        if ( pptx.PresentationPart == null )
            return ps;


        /*
         * 
         */
        ps.Properties = MetaGet( pptx );
        ps.SlideMasters = pptx.PresentationPart.SlideMasterParts.Count();


        /*
         * 
         */
        var presentation = pptx.PresentationPart.Presentation;
        var slideIdList = presentation.SlideIdList;

        if ( slideIdList == null )
            return ps;

        var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();

        ps.Slides = new List<SlideSummary>();

        for ( var i = 0; i < slideIdList.Count(); i++ )
        {
            SlideId slideId = slideIds[ i ];
            var slidePart = (SlidePart) pptx.PresentationPart.GetPartById( slideId.RelationshipId! );

            var ss = new SlideSummary()
            {
                Index = i,
                Title = slidePart.SlideTitle(),
                IsHidden = slidePart.IsHidden(),
                HasNotes = slidePart.NotesSlidePart != null,
            };

            ps.Slides.Add( ss );
        }


        return ps;
    }


    /// <summary />
    public async Task<Result<DocumentRevision>> ReleaseAsync( string filePath )
    {
        await Task.Yield();


        /*
         * 
         */
        var from = Path.Combine( Environment.CurrentDirectory, filePath );
        var dirName = Path.GetDirectoryName( from )!;


        /*
         * 
         */
        var fname = Path.GetFileName( filePath );

        var extn = Path.GetExtension( fname );
        var name = Path.GetFileNameWithoutExtension( fname );

        if ( name.EndsWith( " Live" ) == false )
            return new Result<DocumentRevision>( new ApplicationException( "File must end with 'Live'" ) );

        var rest = name.Substring( 0, name.Length - 5 );


        /*
         * 
         */
        var regex = new Regex( @" Rev(?<rev>\d+)$" );
        var maxRev = 0;

        foreach ( var f in Directory.GetFiles( dirName, "*" + extn ) )
        {
            var fn = Path.GetFileNameWithoutExtension( f );

            if ( fn.StartsWith( rest ) == false )
                continue;

            if ( fn == name )
                continue;

            var m = regex.Match( fn );

            if ( m.Success == false )
                continue;

            var fr = int.Parse( m.Groups[ "rev" ].Value );

            if ( fr > maxRev )
                maxRev = fr;
        }


        /*
         * 
         */
        var rev = new DocumentRevision()
        {
            Number = maxRev + 1,
            Moment = DateTime.Now,
            Filename = "",
        };

        rev.Filename = rest + " " + rev.Moment.ToString( "yyyyMMdd" ) + " Rev" + rev.Number.ToString( "0" ) + extn;


        /*
         * 
         */
        var to = Path.Combine( dirName, rev.Filename );

        File.Copy( from, to, true );
        Mutate( to, rev );

        return new Result<DocumentRevision>( rev );
    }



    /// <summary />
    private void Mutate( string filePath, DocumentRevision? rev = null )
    {
        using var pptx = PresentationDocument.Open( filePath, true );

        if ( pptx.PresentationPart == null )
            return;


        //
        var presentation = pptx.PresentationPart.Presentation;
        var slideIdList = presentation.SlideIdList;

        if ( slideIdList == null )
            return;


        // Check number of slide masters
        //res.SlideMasters = pptx.PresentationPart.SlideMasterParts.Count();


        // Get all slide IDs to process (iterate backwards to safely remove)
        var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();

        for ( int i = slideIds.Count - 1; i >= 0; i-- )
        {
            SlideId slideId = slideIds[ i ];
            var slidePart = (SlidePart) pptx.PresentationPart.GetPartById( slideId.RelationshipId! );


            // Get slide title
            var slideTitle = slidePart.SlideTitle();


            // Remove hidden slides
            if ( slidePart.IsHidden() == true )
            {
                _logger.LogInformation( "#{Number} '{Title}': Hidden, removing slide...", i, slideTitle );

                slideId.Remove();
                pptx.PresentationPart.DeletePart( slidePart );
                continue;
            }


            // Remove modern comments
            var commentParts = slidePart.GetPartsOfType<PowerPointCommentPart>().ToList();

            if ( commentParts.Any() == true )
            {
                _logger.LogInformation( "#{SlideNr} '{SlideTitle}': Has {Count} comment part(s), clearing...", i, slideTitle, commentParts.Count );

                foreach ( var commentPart in commentParts )
                {
                    _logger.LogDebug( "Nr Comments: {NrComments}", commentPart.CommentList.Count() );

                    slidePart.DeletePart( commentPart );
                }
            }


            // Remove legacy comments
            var commentsParts = slidePart.GetPartsOfType<SlideCommentsPart>().ToList();

            if ( commentsParts.Any() == true )
            {
                _logger.LogInformation( "#{SlideNr} '{Title}': Has {Count} comment parts, clearing...", i, slideTitle, commentsParts.Count );

                foreach ( var commentsPart in commentsParts )
                    slidePart.DeletePart( commentsPart );
            }


            // Remove notes from visible slides
            if ( slidePart.NotesSlidePart != null )
            {
                _logger.LogInformation( "#{Number} '{Title}': Has notes, clearing...", i, slideTitle );

                slidePart.DeletePart( slidePart.NotesSlidePart );
            }
        }


        // Remove comment authors (modern)
        var powerPointAuthorsParts = pptx.PresentationPart.GetPartsOfType<PowerPointAuthorsPart>().ToList();

        if ( powerPointAuthorsParts.Any() == true )
        {
            _logger.LogInformation( "Removing (modern) comment authors..." );

            foreach ( var authorsPart in powerPointAuthorsParts )
                pptx.PresentationPart.DeletePart( authorsPart );
        }

        // Remove comment authors (legacy)
        if ( pptx.PresentationPart.CommentAuthorsPart != null )
        {
            _logger.LogInformation( "Removing (legacy) comment authors..." );

            pptx.PresentationPart.DeletePart( pptx.PresentationPart.CommentAuthorsPart );
        }


        // 
        if ( rev != null )
            CoverSet( pptx, rev );

        MetaCleanAuthorAndLastModified( pptx );
        MetaSetRevision( pptx, 1 );


        // 
        presentation.Save();
    }


    /// <summary />
    private void CoverSet( PresentationDocument pptx, DocumentRevision rev )
    {
        if ( pptx.PresentationPart == null )
            return;

        var presentation = pptx.PresentationPart.Presentation;
        var slideIdList = presentation.SlideIdList;

        if ( slideIdList == null || !slideIdList.ChildElements.Any() )
            return;


        // Get the first slide
        var firstSlideId = slideIdList.ChildElements.OfType<SlideId>().FirstOrDefault();

        if ( firstSlideId == null )
            return;

        var firstSlidePart = (SlidePart) pptx.PresentationPart.GetPartById( firstSlideId.RelationshipId! );
        var slide = firstSlidePart.Slide;


        // Get all shapes in the slide
        var shapes = slide.CommonSlideData?.ShapeTree?.Elements<Shape>();

        if ( shapes == null )
            return;


        foreach ( var shape in shapes )
        {
            var textBody = shape.TextBody;

            if ( textBody == null )
                continue;

            var paragraphs = textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>();

            foreach ( var paragraph in paragraphs )
            {
                var runs = paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>().ToList();

                foreach ( var run in runs )
                {
                    if ( run.Text == null )
                        continue;

                    var text = run.Text.Text?.Trim();

                    if ( text == "LIVE" )
                    {
                        string revisionText = $"Rev {rev.Number}";

                        _logger.LogInformation( "Updating revision text from 'LIVE' to '{Revision}'", revisionText );
                        run.Text.Text = revisionText;
                    }

                    if ( text == "<DATE>" )
                    {
                        string date = rev.Moment.ToString( "dd MMMM yyyy" );

                        _logger.LogInformation( "Updating date from '<DATE>' to '{Date}'", date );
                        run.Text.Text = date;
                    }
                }
            }
        }
    }



    /// <summary />
    private PresentationProps MetaGet( PresentationDocument pptx )
    {
        var coreProperties = pptx.PackageProperties;

        return new PresentationProps()
        {
            Creator = coreProperties.Creator?.Nullify(),
            LastModifiedBy = coreProperties.LastModifiedBy?.Nullify(),
            Title = coreProperties.Title?.Nullify(),
            Subject = coreProperties.Subject?.Nullify(),
            Keywords = coreProperties.Keywords?.Nullify(),
            Description = coreProperties.Description?.Nullify(),
            Version = coreProperties.Version?.Nullify(),
        };
    }


    /// <summary />
    private void MetaCleanAuthorAndLastModified( PresentationDocument pptx )
    {
        var coreProperties = pptx.PackageProperties;

        if ( string.IsNullOrEmpty( coreProperties.Creator ) == false )
        {
            _logger.LogInformation( "Clearing creator: {Creator}", coreProperties.Creator );
            coreProperties.Creator = string.Empty;
        }

        if ( string.IsNullOrEmpty( coreProperties.LastModifiedBy ) == false )
        {
            _logger.LogInformation( "Clearing last modified by: {LastModifiedBy}", coreProperties.LastModifiedBy );
            coreProperties.LastModifiedBy = string.Empty;
        }
    }


    /// <summary />
    private void MetaSetRevision( PresentationDocument pptx, int revision )
    {
        var coreProperties = pptx.PackageProperties;
        var rev = "Rev " + revision;

        _logger.LogInformation( "Powerpoint version: {Version}", rev );
        coreProperties.Version = rev;
    }


    /// <summary />
    private void MetaClean( PresentationDocument pptx )
    {
        var coreProperties = pptx.PackageProperties;

        // Optionally clear other metadata
        if ( !string.IsNullOrEmpty( coreProperties.Title ) )
        {
            _logger.LogInformation( "Clearing title: {Title}", coreProperties.Title );
            coreProperties.Title = string.Empty;
        }

        if ( !string.IsNullOrEmpty( coreProperties.Subject ) )
        {
            _logger.LogInformation( "Clearing subject: {Subject}", coreProperties.Subject );
            coreProperties.Subject = string.Empty;
        }

        if ( !string.IsNullOrEmpty( coreProperties.Keywords ) )
        {
            _logger.LogInformation( "Clearing keywords: {Keywords}", coreProperties.Keywords );
            coreProperties.Keywords = string.Empty;
        }

        if ( !string.IsNullOrEmpty( coreProperties.Description ) )
        {
            _logger.LogInformation( "Clearing description" );
            coreProperties.Description = string.Empty;
        }

        if ( !string.IsNullOrEmpty( coreProperties.Category ) )
        {
            _logger.LogInformation( "Clearing category: {Category}", coreProperties.Category );
            coreProperties.Category = string.Empty;
        }
    }
}