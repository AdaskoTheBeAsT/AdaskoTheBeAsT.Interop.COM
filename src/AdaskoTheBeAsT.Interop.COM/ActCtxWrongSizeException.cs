using System;
#if NETSTANDARD2_0
using System.Runtime.Serialization;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

[Serializable]
#if NET8_0_OR_GREATER
#pragma warning disable S3925 // "ISerializable" should be implemented correctly
#endif
public class ActCtxWrongSizeException
#if NET8_0_OR_GREATER
#pragma warning restore S3925 // "ISerializable" should be implemented correctly
#endif
    : Exception
{
    public ActCtxWrongSizeException()
    {
    }

    public ActCtxWrongSizeException(string message)
        : base(message)
    {
    }

    public ActCtxWrongSizeException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

#if NETSTANDARD2_0
    protected ActCtxWrongSizeException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
#endif
}
