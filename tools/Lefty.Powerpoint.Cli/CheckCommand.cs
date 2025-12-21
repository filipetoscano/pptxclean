using McMaster.Extensions.CommandLineUtils;
using System.ComponentModel.DataAnnotations;

namespace Lefty.Powerpoint.Cli;

/// <summary />
[Command( "check", Description = "Validates a PowerPoint file" )]
public class CheckCommand
{
    private readonly PowerpointService _svc;


    /// <summary />
    public CheckCommand( PowerpointService svc )
    {
        _svc = svc;
    }


    /// <summary />
    [Argument( 0, Description = "Live version of powerpoint file" )]
    [Required]
    [FileExists]
    public string? InputFile { get; set; }


    /// <summary />
    public async Task<int> OnExecuteAsync()
    {
        await _svc.CheckAsync( this.InputFile! );

        return 0;
    }
}