using McMaster.Extensions.CommandLineUtils;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace Lefty.Powerpoint.Cli;

/// <summary />
[Command( "pptxdo" )]
[Subcommand( typeof( CleanCommand ))]
public class Program
{
    /// <summary />
    public static int Main( string[] args )
    {
        /*
         * 
         */
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.Console()
            .CreateLogger();

        var logger = Log.ForContext<Program>();


        /*
         * 
         */
        var svc = new ServiceCollection();

        svc.AddLogging( loggingBuilder =>
            loggingBuilder.AddSerilog( dispose: true ) );

        svc.AddTransient<PowerpointCleanerService>();

        var sp = svc.BuildServiceProvider();


        /*
         * 
         */
        var app = new CommandLineApplication<Program>();

        try
        {
            app.Conventions
                .UseDefaultConventions()
                .UseConstructorInjection( sp );
        }
        catch ( Exception ex )
        {
            logger.Fatal( ex, "Unhandled exception" );

            return 2;
        }


        /*
         * 
         */
        try
        {
            return app.Execute( args );
        }
        catch ( UnrecognizedCommandParsingException ex )
        {
            logger.Error( ex, "Unrecognized command" );

            return 2;
        }
        catch ( Exception ex )
        {
            logger.Fatal( ex, "Unhandled exception" );

            return 2;
        }
    }
}