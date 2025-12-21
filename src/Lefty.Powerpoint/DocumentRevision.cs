namespace Lefty.Powerpoint;

/// <summary />
public class DocumentRevision
{
    /// <summary />
    public required int Number { get; set; }

    /// <summary />
    public required DateTime Moment { get; set; }

    /// <summary />
    public required string Filename { get; set; }
}