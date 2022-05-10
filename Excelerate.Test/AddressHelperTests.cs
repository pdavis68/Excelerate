using System.Drawing;
using Xunit;
using Excelerate;

namespace Excelerate.Test;

public class AddressHelperTests
{
    [Fact]
    public void TestToRowCol()
    {
        var cell = AddressHelper.ToRowCol("A1");
        Assert.Equal(new Point(1, 1), cell);
        cell = AddressHelper.ToRowCol("AA1");
        Assert.Equal(new Point(27, 1), cell);
        cell = AddressHelper.ToRowCol("AC1024");
        Assert.Equal(new Point(29, 1024), cell);
        cell = AddressHelper.ToRowCol("AC1024");
        Assert.Equal(new Point(29, 1024), cell);
        cell = AddressHelper.ToRowCol("LL324");
        Assert.Equal(new Point(324, 324), cell);
        cell = AddressHelper.ToRowCol("XA626");
        Assert.Equal(new Point(625, 626), cell);
    }

    [Fact]
    public void TestToExcelCoordinates()
    {
        var cell = AddressHelper.ToExcelCoordinates(1, 1);
        Assert.Equal("A1", cell);
        cell = AddressHelper.ToExcelCoordinates(27, 1);
        Assert.Equal("AA1", cell);
        cell = AddressHelper.ToExcelCoordinates(29, 1024);
        Assert.Equal("AC1024", cell);
        cell = AddressHelper.ToExcelCoordinates(29, 1024);
        Assert.Equal("AC1024", cell);
        cell = AddressHelper.ToExcelCoordinates(324, 324);
        Assert.Equal("LL324", cell);
        cell = AddressHelper.ToExcelCoordinates(625, 626);
        Assert.Equal("XA626", cell);
    }    
}