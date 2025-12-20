using McMaster.Extensions.CommandLineUtils;

namespace Lefty.Powerpoint.Cli;

/// <summary />
[Command( "clean" )]
public class CleanCommand
{
    private readonly PowerpointCleanerService _svc;


    /// <summary />
    public CleanCommand( PowerpointCleanerService svc )
    {
        _svc = svc;
    }


    /// <summary />
    public async Task<int> OnExecuteAsync()
    {
        await _svc.CleanAsync( "Presentation1.pptx", "Presentation2.pptx" );

        return 0;
    }
}