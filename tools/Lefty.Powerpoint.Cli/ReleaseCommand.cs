using McMaster.Extensions.CommandLineUtils;
using Microsoft.Extensions.Logging;
using System.ComponentModel.DataAnnotations;

namespace Lefty.Powerpoint.Cli;

/// <summary />
[Command( "release", Description = "Creates a revision of Powerpoint file: cleans and increments revision" )]
public class ReleaseCommand
{
    private readonly PowerpointService _svc;
    private readonly ILogger<ReleaseCommand> _logger;


    /// <summary />
    public ReleaseCommand( PowerpointService svc, ILogger<ReleaseCommand> logger )
    {
        _svc = svc;
        _logger = logger;
    }


    /// <summary />
    [Argument( 0, Description = "Live version of powerpoint file" )]
    [Required]
    [FileExists]
    public string? InputFile { get; set; }


    /// <summary />
    public async Task<int> OnExecuteAsync()
    {
        var res = await _svc.ReleaseAsync( this.InputFile! );

        if ( res.IsOk == false )
        {
            _logger.LogError( res.Exception.Message );
            return 1;
        }

        _logger.LogInformation( "Released: {Revision}, written to {Filename}", res.Value.Number, res.Value.Filename );
        return 0;
    }
}