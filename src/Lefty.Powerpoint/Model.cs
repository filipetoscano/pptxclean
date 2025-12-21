namespace Lefty.Powerpoint;

/// <summary />
public class PresentationSummary
{
    /// <summary />
    public PresentationProps? Properties { get; set; }

    /// <summary />
    public List<SlideSummary>? Slides { get; set; }

    /// <summary />
    public int SlideMasters { get; set; }
}


/// <summary />
public class PresentationProps
{
    /// <summary />
    public required string? Title { get; set; }

    /// <summary />
    public required string? Creator { get; set; }

    /// <summary />
    public required string? LastModifiedBy { get; set; }

    /// <summary />
    public required string? Subject { get; set; }

    /// <summary />
    public required string? Keywords { get; set; }

    /// <summary />
    public required string? Description { get; set; }

    /// <summary />
    public required string? Version { get; set; }
}


/// <summary />
public class SlideSummary
{
    /// <summary />
    public required int Index { get; set; }

    /// <summary />
    public required string? Title { get; set; }

    /// <summary />
    public required bool IsHidden { get; set; }

    /// <summary />
    public bool HasNotes { get; set; }

    /// <summary />
    public bool HasComments { get; set; }

    /// <summary />
    public List<SlideComment>? Comments { get; set; }
}


/// <summary />
public class SlideComment
{
    /// <summary />
    public string? Author { get; set; }

    /// <summary />
    public string? Text { get; set; }
}