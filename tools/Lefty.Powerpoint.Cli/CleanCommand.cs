using McMaster.Extensions.CommandLineUtils;
using Microsoft.Extensions.Logging;
using System.ComponentModel.DataAnnotations;

namespace Lefty.Powerpoint.Cli;

/// <summary />
[Command( "clean" )]
public class CleanCommand
{
    private readonly PowerpointService _svc;
    private readonly ILogger<CleanCommand> _logger;


    /// <summary />
    public CleanCommand( PowerpointService svc, ILogger<CleanCommand> logger )
    {
        _svc = svc;
        _logger = logger;
    }


    /// <summary />
    [Argument( 0, Description = "Powerpoint file" )]
    [Required]
    [FileExists]
    public string? InputFile { get; set; }

    /// <summary />
    [Option( "-a|--auto", CommandOptionType.NoValue, Description = "Auto naming of output file" )]
    public bool UseAuto { get; set; }

    /// <summary />
    [Option( "-f|--force", CommandOptionType.NoValue, Description = "In place cleaning of file [!! DESTRUCTIVE !!]" )]
    public bool UseForce { get; set; }

    /// <summary />
    [Option( "-o|--output", CommandOptionType.SingleValue, Description = "Output file" )]
    [FileNotExists]
    public string? OutputFile { get; set; }


    /// <summary />
    public async Task<int> OnExecuteAsync()
    {
        if ( this.OutputFile != null && this.UseForce == true )
        {
            _logger.LogError( "Force and Output files are mutually exclusive options" );
            return 1;
        }

        if ( this.OutputFile != null && this.UseAuto == true )
        {
            _logger.LogError( "Auto and Output files are mutually exclusive options" );
            return 1;
        }

        if ( this.UseAuto == true && this.UseForce == true )
        {
            _logger.LogError( "Auto and Force are mutually exclusive options" );
            return 1;
        }

        if ( this.OutputFile == null && this.UseAuto == false && this.UseForce == false )
        {
            _logger.LogError( "One of the following options must be set: Auti, Force, Output" );
            return 1;
        }



        /*
         * 
         */
        string file;

        if ( this.OutputFile != null )
        {
            File.Copy( this.InputFile!, this.OutputFile );
            file = this.OutputFile;
        }
        else if ( this.UseForce == true )
        {
            file = this.InputFile!;
        }
        else if ( this.UseAuto == true )
        {
            throw new NotImplementedException();
        }
        else
        {
            throw new NotSupportedException();
        }

        await _svc.CleanAsync( file );

        return 0;
    }
}