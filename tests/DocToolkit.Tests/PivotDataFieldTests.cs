using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class PivotDataFieldTests
{
    [Fact]
    public void Constructs_WithFieldNameAndFunction()
    {
        var field = new PivotDataField("Amount", PivotFunction.Sum);

        Assert.Equal("Amount", field.FieldName);
        Assert.Equal(PivotFunction.Sum, field.Function);
    }

    [Fact]
    public void NullFieldName_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => new PivotDataField(null!, PivotFunction.Sum));
    }

    [Fact]
    public void BlankFieldName_Throws()
    {
        Assert.Throws<ArgumentException>(() => new PivotDataField("   ", PivotFunction.Sum));
    }
}
