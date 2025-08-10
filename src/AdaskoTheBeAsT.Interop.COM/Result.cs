using System;

namespace AdaskoTheBeAsT.Interop.COM;

public sealed class Result
{
    public Exception? Exception { get; set; }

    public bool Success { get; set; }
}
