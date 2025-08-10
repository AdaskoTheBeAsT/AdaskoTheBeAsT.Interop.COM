using System.Reflection;
using AwesomeAssertions;
using Xunit;

namespace AdaskoTheBeAsT.Interop.COM.Test;

public class ExecutorTest
{
    [Fact]
    public void AddStringsTest()
    {
        // Arrange
        var currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
        var comAssemblyPath = Path.Combine(currentPath, @"..\..\..\..\..\..\x86\Debug\NativeCOM.dll");
        var manifestPath = Path.Combine(currentPath, @"..\..\..\..\..\..\Manifest\NativeCOM\NativeCOM.manifest");

        File.Exists(comAssemblyPath).Should().BeTrue();
        File.Exists(manifestPath).Should().BeTrue();
        var str1 = "Hello";
        var str2 = "World!";
        string? output = null;
        Action action = () =>
        {
            var executor = new NativeCOM.StringConcatenatorClass();
            output = executor.ConcatStrings(str1, str2);
        };

        // Act
        var result = Executor.Execute(comAssemblyPath, manifestPath, action);

        // Assert
        result.Should().NotBeNull();
        result.Success.Should().BeTrue();
    }
}
