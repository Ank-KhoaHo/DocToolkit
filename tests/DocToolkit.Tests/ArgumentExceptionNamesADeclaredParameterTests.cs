using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// An <see cref="ArgumentException"/> must name a parameter the invoked method actually declares.
/// </summary>
/// <remarks>
/// <b>DERIVED FROM THE PUBLIC SURFACE, not from a list.</b> Two instances of this defect were found
/// by review on 2026-08-20, and both sat next to a comment describing the correct pattern - the
/// pattern was written down and then not applied to the neighbouring overloads. A third copy of
/// that list would go stale the same way, so this walks the shipped types by reflection and a new
/// file-path overload is covered the moment it exists.
///
/// <para><b>The input is an EMPTY FILE, and that is the whole trick.</b> A null or whitespace path
/// is rejected by the overload's own guard, which names the right parameter - so those cases pass
/// whether or not the bug is present. An empty file gets past that guard and reaches the byte[] or
/// Stream implementation underneath, which raises the exception naming ITS OWN parameter. That is
/// the caller being told to check an argument they never passed.</para>
/// </remarks>
public class ArgumentExceptionNamesADeclaredParameterTests
{
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    /// <summary>A plausible value for every parameter that is not the one under test.</summary>
    private static object? Placeholder(ParameterInfo p, string emptyFile, string outFile)
    {
        var t = p.ParameterType;

        if (t == typeof(CancellationToken)) return CancellationToken.None;
        if (t == typeof(string))
        {
            return p.Name switch
            {
                "outputPath" => outFile,
                "inputPath" or "path" => emptyFile,
                "sheetName" => "Sheet1",
                "cellRef" => "A1",
                "collection" => "items",
                "placeholder" => "{{x}}",
                "html" => "<p>x</p>",
                _ => "x",
            };
        }
        if (t == typeof(byte[])) return TinyPng;
        if (t == typeof(bool)) return false;
        if (t == typeof(int)) return 0;
        if (t == typeof(PageSetup)) return PageSetup.A4;
        if (t == typeof(double?) || Nullable.GetUnderlyingType(t) is not null) return null;
        if (t == typeof(IReadOnlyDictionary<string, string>))
            return new Dictionary<string, string> { ["{{x}}"] = "y" };
        if (t.IsGenericType && typeof(System.Collections.IEnumerable).IsAssignableFrom(t))
            return Activator.CreateInstance(typeof(List<>).MakeGenericType(t.GetGenericArguments()[0]));
        if (t == typeof(object)) return "x";
        return t.IsValueType ? Activator.CreateInstance(t) : null;
    }

    [Fact]
    public async Task EveryFilePathOverload_NamesAParameterItDeclares()
    {
        var types = typeof(DocxEditor).Assembly.GetExportedTypes()
            .Where(t => t.IsAbstract && t.IsSealed)   // static classes
            .OrderBy(t => t.Name);

        var dir = Directory.CreateTempSubdirectory("doctoolkit-paramname");
        var emptyFile = Path.Join(dir.FullName, "empty.bin");
        var outFile = Path.Join(dir.FullName, "out.bin");
        await File.WriteAllBytesAsync(emptyFile, Array.Empty<byte>());

        var exercised = 0;
        var observed = 0;
        var wrong = new List<string>();

        try
        {
            foreach (var type in types)
                foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Static)
                             .OrderBy(m => m.Name).ThenBy(m => m.GetParameters().Length))
                {
                    var parameters = method.GetParameters();
                    var victim = parameters.FirstOrDefault(
                        p => p.ParameterType == typeof(string) && p.Name is "path" or "inputPath");
                    if (victim is null) continue;

                    object?[] args;
                    try
                    {
                        args = parameters.Select(p => Placeholder(p, emptyFile, outFile)).ToArray();
                    }
                    catch
                    {
                        continue;   // cannot build a call for this shape
                    }

                    exercised++;
                    var declared = parameters.Select(p => p.Name).ToHashSet(StringComparer.Ordinal);

                    try
                    {
                        var returned = method.Invoke(null, args);
                        if (returned is Task task) await task;
                    }
                    catch (Exception ex)
                    {
                        var real = ex is TargetInvocationException tie ? tie.InnerException! : ex;
                        if (real is not ArgumentException arg) continue;

                        observed++;
                        if (arg.ParamName is null || !declared.Contains(arg.ParamName))
                        {
                            wrong.Add($"{type.Name}.{method.Name}({string.Join(", ", declared)}) "
                                + $"threw ArgumentException naming '{arg.ParamName}', which it does not declare");
                        }
                    }
                }
        }
        finally
        {
            dir.Delete(recursive: true);
        }

        // NOT VACUOUS. A walk that invoked nothing, or that never provoked the exception this test
        // is about, proves nothing and must fail rather than pass quietly - the standard the guard
        // scripts in this repository are already held to.
        Assert.True(exercised >= 20, $"only {exercised} file-path overloads were exercised");
        Assert.True(observed >= 1,
            $"{exercised} overloads ran and not one raised an ArgumentException, so the assertion "
            + "under test never executed");

        Assert.True(wrong.Count == 0, string.Join("\n", wrong));
    }
}
