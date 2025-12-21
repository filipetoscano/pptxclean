namespace Lefty.Powerpoint;

/// <summary />
public struct Result<T>
{
    private readonly bool _ok;
    private readonly T? _value;
    private readonly Exception? _exception;


    /// <summary />
    public Result( T value )
    {
        _ok = true;
        _value = value;
    }


    /// <summary />
    public Result( Exception exception )
    {
        _ok = false;
        _exception = exception;
    }


    /// <summary />
    public bool IsOk { get => _ok ; }


    /// <summary />
    public T Value
    {
        get
        {
            if ( _ok == false )
                throw new InvalidOperationException();

            if ( _value == null )
                throw new InvalidProgramException();

            return _value;
        }
    }


    /// <summary />
    public Exception Exception
    {
        get
        {
            if ( _ok == true )
                throw new InvalidOperationException();

            if ( _exception == null )
                throw new InvalidProgramException();

            return _exception;
        }
    }
}