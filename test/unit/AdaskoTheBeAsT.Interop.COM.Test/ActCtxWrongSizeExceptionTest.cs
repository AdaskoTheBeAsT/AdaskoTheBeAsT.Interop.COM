using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public class ActCtxWrongSizeExceptionTest
{
    [Fact]
    public void DefaultConstructorShouldProduceInstance()
    {
        var sut = new ActCtxWrongSizeException();

        sut.Should().BeOfType<ActCtxWrongSizeException>();
        sut.InnerException.Should().BeNull();
    }

    [Fact]
    public void MessageConstructorShouldSetMessage()
    {
        const string message = "wrong size";

        var sut = new ActCtxWrongSizeException(message);

        sut.Message.Should().Be(message);
    }

    [Fact]
    public void InnerExceptionConstructorShouldSetMessageAndInnerException()
    {
        const string message = "wrong size";
        var inner = new InvalidOperationException("inner");

        var sut = new ActCtxWrongSizeException(message, inner);

        sut.Message.Should().Be(message);
        sut.InnerException.Should().BeSameAs(inner);
    }
}
