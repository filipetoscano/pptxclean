namespace Lefty.Powerpoint;

/// <summary />
internal static class SystemExtensions
{
    /// <summary />
    internal static string? Nullify( this string? value )
    {
        if ( value == null )
            return null;

        if ( value.Length == 0 )
            return null;

        return value;
    }
}