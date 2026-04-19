using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public class ComExecutorTest
{
    [Fact]
    public void ShouldImplementIComExecutor()
    {
        var sut = new ComExecutor();

        sut.Should().BeAssignableTo<IComExecutor>();
    }

    [Fact]
    public void ExecuteShouldDelegateToStaticExecutor()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();
        IComExecutor sut = new ComExecutor();
        string? output = null;

        var result = sut.Execute(comAssemblyPath, manifestPath, () =>
        {
            var obj = new NativeCOM.StringConcatenatorClass();
            output = obj.ConcatStrings("Hello", "World!");
        });

        result.Success.Should().BeTrue();
        output.Should().Be("HelloWorld!");
    }

    [Fact]
    public void CreateAndFreeShouldDelegateToStaticExecutor()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();
        IComExecutor sut = new ComExecutor();

        var creation = sut.Create(
            comAssemblyPath,
            manifestPath,
            () => new NativeCOM.StringConcatenatorClass());

        creation.Success.Should().BeTrue();
        creation.Value.Should().NotBeNull();

        var release = sut.Free(creation.Value!);

        release.Success.Should().BeTrue();
        creation.Value!.IsReleased.Should().BeTrue();
    }

    [Fact]
    public void ExecuteCollectionOverloadShouldDelegateToStaticExecutor()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();
        var descriptors = new[] { new ComPathDescriptor(comAssemblyPath, manifestPath) };
        IComExecutor sut = new ComExecutor();
        string? output = null;

        var result = sut.Execute(descriptors, () =>
        {
            var obj = new NativeCOM.StringConcatenatorClass();
            output = obj.ConcatStrings("Quick", "Brown");
        });

        result.Success.Should().BeTrue();
        output.Should().Be("QuickBrown");
    }

    [Fact]
    public void CreateCollectionOverloadShouldDelegateToStaticExecutor()
    {
        var (comAssemblyPath, manifestPath) = GetPaths();
        var descriptors = new[] { new ComPathDescriptor(comAssemblyPath, manifestPath) };
        IComExecutor sut = new ComExecutor();

        var creation = sut.Create(
            descriptors,
            () => new NativeCOM.StringConcatenatorClass());

        creation.Success.Should().BeTrue();
        creation.Value.Should().NotBeNull();

        sut.Free(creation.Value!).Success.Should().BeTrue();
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
