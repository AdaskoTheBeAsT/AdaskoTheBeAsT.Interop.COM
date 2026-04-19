using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public class ComPathDescriptorTest
{
    [Fact]
    public void ConstructorShouldAssignProperties()
    {
        const string asm = @"C:\temp\MyCom.dll";
        const string manifest = @"C:\temp\MyCom.manifest";

        var sut = new ComPathDescriptor(asm, manifest);

        sut.ComAssemblyPath.Should().Be(asm);
        sut.ComManifestPath.Should().Be(manifest);
    }

    [Fact]
    public void ConstructorShouldThrowWhenAssemblyPathIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => new ComPathDescriptor(null!, @"C:\temp\MyCom.manifest"));

        ex.ParamName.Should().Be("comAssemblyPath");
    }

    [Fact]
    public void ConstructorShouldThrowWhenManifestPathIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => new ComPathDescriptor(@"C:\temp\MyCom.dll", null!));

        ex.ParamName.Should().Be("comManifestPath");
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("\t")]
    public void ConstructorShouldThrowWhenAssemblyPathIsEmptyOrWhitespace(string value)
    {
        var ex = Assert.Throws<ArgumentException>(
            () => new ComPathDescriptor(value, @"C:\temp\MyCom.manifest"));

        ex.ParamName.Should().Be("comAssemblyPath");
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("\t")]
    public void ConstructorShouldThrowWhenManifestPathIsEmptyOrWhitespace(string value)
    {
        var ex = Assert.Throws<ArgumentException>(
            () => new ComPathDescriptor(@"C:\temp\MyCom.dll", value));

        ex.ParamName.Should().Be("comManifestPath");
    }
}
