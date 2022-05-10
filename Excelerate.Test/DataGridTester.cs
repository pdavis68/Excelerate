using System;
using System.Diagnostics;
using Excelerate;
using Xunit;

namespace Excelerate.Test
{
    public class DataGridTester
    {

        [Fact]
        public void Benchmark1()
        {
            int numRows = 500000;
            int numCols = 20;
            var sw = new Stopwatch();
            var dg = new DataGrid();
            Debug.WriteLine($"Populating grid with {numRows} rows and {numCols} cols");
            sw.Start();
            for(int rowIndex = 1; rowIndex < numRows; rowIndex++)
            {
                for(int colIndex = 1; colIndex < numCols; colIndex++)
                {
                    dg.InsertItem(rowIndex, colIndex, GetValue(10, 30));
                }
            }
            sw.Stop();
            Debug.WriteLine($"Insert Time: {sw.ElapsedMilliseconds} ms");

            Debug.WriteLine($"Retrieving all rows");
            sw.Start();
            for(int rowIndex = 1; rowIndex < numRows; rowIndex++)
            {
                dg.GetRow(rowIndex);
            }
            Debug.WriteLine($"Retrieval Time: {sw.ElapsedMilliseconds} ms");
        }

        private Random _rand = new Random();
        private string _chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789)!@#$%^%&*(";
        private string GetValue(int baseVal, int variable)
        {
            int count = baseVal + _rand.Next(variable);
            var str = String.Empty;
            while(count-- > 0)
            {
                str += _chars[_rand.Next(_chars.Length)];
            }
            return str;
        }
    }
}