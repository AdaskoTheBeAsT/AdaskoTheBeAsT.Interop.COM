using System;
#if NETSTANDARD2_0
using System.Runtime.Serialization;
#endif

namespace AdaskoTheBeAsT.Interop.COM;

/// <summary>
/// The exception that is thrown when the native activation-context structure size does not match the current process architecture.
/// </summary>
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
    /// <summary>
    /// Initializes a new instance of the <see cref="ActCtxWrongSizeException"/> class.
    /// </summary>
    public ActCtxWrongSizeException()
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ActCtxWrongSizeException"/> class with a specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    public ActCtxWrongSizeException(string message)
        : base(message)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ActCtxWrongSizeException"/> class with a specified error message and a reference to the inner exception.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="innerException">The exception that caused the current exception.</param>
    public ActCtxWrongSizeException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

#if NETSTANDARD2_0
    /// <summary>
    /// Initializes a new instance of the <see cref="ActCtxWrongSizeException"/> class with serialized data.
    /// </summary>
    /// <param name="info">The serialization information that holds the serialized object data.</param>
    /// <param name="context">The contextual information about the source or destination.</param>
    protected ActCtxWrongSizeException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
#endif
}
