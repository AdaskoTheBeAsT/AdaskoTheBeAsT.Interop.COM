using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public class ExecuteValidationTest
{
    [Fact]
    public void ExecuteShouldThrowWhenAssemblyPathIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => Executor.Execute(null!, "manifest", () => { }));

        ex.ParamName.Should().Be("comAssemblyPath");
    }

    [Fact]
    public void ExecuteShouldThrowWhenManifestPathIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => Executor.Execute("assembly", null!, () => { }));

        ex.ParamName.Should().Be("manifestPath");
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    public void ExecuteShouldThrowWhenAssemblyPathIsWhitespace(string value)
    {
        var ex = Assert.Throws<ArgumentException>(
            () => Executor.Execute(value, "manifest", () => { }));

        ex.ParamName.Should().Be("comAssemblyPath");
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    public void ExecuteShouldThrowWhenManifestPathIsWhitespace(string value)
    {
        var ex = Assert.Throws<ArgumentException>(
            () => Executor.Execute("assembly", value, () => { }));

        ex.ParamName.Should().Be("manifestPath");
    }

    [Fact]
    public void ExecuteCollectionOverloadShouldThrowWhenDescriptorsIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => Executor.Execute((ICollection<ComPathDescriptor>)null!, () => { }));

        ex.ParamName.Should().Be("comPathDescriptors");
    }

    [Fact]
    public void ExecuteCollectionOverloadShouldThrowWhenDescriptorsIsEmpty()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => Executor.Execute(Array.Empty<ComPathDescriptor>(), () => { }));

        ex.ParamName.Should().Be("comPathDescriptors");
    }

    [Fact]
    public void CreateCollectionOverloadShouldThrowWhenDescriptorsIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => Executor.Create<object>((ICollection<ComPathDescriptor>)null!, () => new object()));

        ex.ParamName.Should().Be("comPathDescriptors");
    }

    [Fact]
    public void CreateCollectionOverloadShouldThrowWhenDescriptorsIsEmpty()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => Executor.Create<object>(Array.Empty<ComPathDescriptor>(), () => new object()));

        ex.ParamName.Should().Be("comPathDescriptors");
    }

    [Fact]
    public void ExecuteShouldReturnFailureResultWhenActionThrows()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();
        var boom = new InvalidOperationException("boom");

        var result = Executor.Execute(
            comAssemblyPath,
            manifestPath,
            () => throw boom);

        result.Success.Should().BeFalse();
        result.Exception.Should().BeSameAs(boom);
    }

    [Fact]
    public void ExecuteCollectionOverloadShouldRunAction()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();
        var descriptors = new[] { new ComPathDescriptor(comAssemblyPath, manifestPath) };
        string? output = null;

        var result = Executor.Execute(descriptors, () =>
        {
            var obj = new NativeCOM.StringConcatenatorClass();
            output = obj.ConcatStrings("Foo", "Bar");
        });

        result.Success.Should().BeTrue();
        result.Exception.Should().BeNull();
        output.Should().Be("FooBar");
    }

    [Fact]
    public void ExecuteCollectionOverloadShouldReturnFailureResultWhenActionThrows()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();
        var descriptors = new[] { new ComPathDescriptor(comAssemblyPath, manifestPath) };
        var boom = new InvalidOperationException("boom-collection");

        var result = Executor.Execute(
            descriptors,
            () => throw boom);

        result.Success.Should().BeFalse();
        result.Exception.Should().BeSameAs(boom);
    }

    [Fact]
    public void CreateShouldThrowWhenAssemblyPathIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => Executor.Create<object>(null!, "manifest", () => new object()));

        ex.ParamName.Should().Be("comAssemblyPath");
    }

    [Fact]
    public void CreateShouldThrowWhenManifestPathIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => Executor.Create<object>("assembly", null!, () => new object()));

        ex.ParamName.Should().Be("manifestPath");
    }

    [Fact]
    public void CreateShouldThrowWhenFactoryIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => Executor.Create<object>("assembly", "manifest", null!));

        ex.ParamName.Should().Be("factory");
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    public void CreateShouldThrowWhenAssemblyPathIsWhitespace(string value)
    {
        var ex = Assert.Throws<ArgumentException>(
            () => Executor.Create<object>(value, "manifest", () => new object()));

        ex.ParamName.Should().Be("comAssemblyPath");
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    public void CreateShouldThrowWhenManifestPathIsWhitespace(string value)
    {
        var ex = Assert.Throws<ArgumentException>(
            () => Executor.Create<object>("assembly", value, () => new object()));

        ex.ParamName.Should().Be("manifestPath");
    }

    [Fact]
    public void CreateCollectionOverloadShouldThrowWhenFactoryIsNull()
    {
        var descriptors = new[] { new ComPathDescriptor("dll", "manifest") };

        var ex = Assert.Throws<ArgumentNullException>(
            () => Executor.Create<object>(descriptors, null!));

        ex.ParamName.Should().Be("factory");
    }

    [Fact]
    public void CreateShouldReturnFailureResultWhenFactoryThrows()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();
        var boom = new InvalidOperationException("factory-boom");

        var creation = Executor.Create<object>(
            comAssemblyPath,
            manifestPath,
            () => throw boom);

        creation.Success.Should().BeFalse();
        creation.Value.Should().BeNull();
        creation.Exception.Should().BeSameAs(boom);
    }

    [Fact]
    public void CreateShouldReturnFailureResultWhenFactoryReturnsNull()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();

        var creation = Executor.Create<object>(
            comAssemblyPath,
            manifestPath,
            () => null!);

        creation.Success.Should().BeFalse();
        creation.Value.Should().BeNull();
        creation.Exception.Should().BeOfType<InvalidOperationException>();
    }

    [Fact]
    public void FreeShouldThrowWhenHandleIsNull()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => Executor.Free<object>(null!));

        ex.ParamName.Should().Be("comObjectHandle");
    }

    [Fact]
    public void FreeShouldReturnSuccessWhenHandleAlreadyReleased()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();
        var creation = Executor.Create(
            comAssemblyPath,
            manifestPath,
            () => new NativeCOM.StringConcatenatorClass());

        Executor.Free(creation.Value!).Success.Should().BeTrue();
        var secondRelease = Executor.Free(creation.Value!);

        secondRelease.Success.Should().BeTrue();
    }

    private static (string ComAssemblyPath, string ManifestPath) GetPaths()
    {
        var currentPath = AppContext.BaseDirectory;
        var comAssemblyPath = Path.GetFullPath(Path.Combine(currentPath, "NativeCOM.dll"));
        var manifestPath = Path.GetFullPath(Path.Combine(currentPath, "NativeCOM.manifest"));

        File.Exists(comAssemblyPath).Should().BeTrue();
        File.Exists(manifestPath).Should().BeTrue();

        return (comAssemblyPath, manifestPath);
    }
}
