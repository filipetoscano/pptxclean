using McMaster.Extensions.CommandLineUtils;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;

namespace Lefty.Powerpoint.Cli;

/// <summary />
[Command( "get", Description = "Retrieves the content of Powerpoint file as JSON" )]
public class GetCommand
{
    private readonly PowerpointService _svc;


    /// <summary />
    public GetCommand( PowerpointService svc )
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
        var m = await _svc.GetAsync( this.InputFile! );

        var jso = new JsonSerializerOptions() { WriteIndented = true };
        var json = JsonSerializer.Serialize( m, jso );

        Console.WriteLine( json );

        return 0;
    }
}