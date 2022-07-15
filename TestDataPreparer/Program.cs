using DataPreparer;
using System;
using Microsoft.Office.Interop.Excel;
using Microsoft.Data.Analysis;
using DF = Microsoft.Data.Analysis.DataFrame;

namespace TestDataPreparer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //var manager = new WbManager();
            //Console.WriteLine(manager.GetWorkbook(@"C:\Users\Why\Documents\Study\Stock\0700.xlsx").Name);
            PrimitiveDataFrameColumn<DateTime> dateTimes = new PrimitiveDataFrameColumn<DateTime>("DateTimes"); // Default length is 0.
            PrimitiveDataFrameColumn<int> ints = new PrimitiveDataFrameColumn<int>("Ints", 3); // Makes a column of length 3. Filled with nulls initially
            StringDataFrameColumn strings = new StringDataFrameColumn("Strings", 3);// Makes a column of length 3. Filled with nulls initially
            dateTimes.Append(DateTime.Parse("2019/01/01"));
            dateTimes.Append(DateTime.Parse("2019/01/01"));
            dateTimes.Append(DateTime.Parse("2019/01/02"));
            DataPreparer.DataFrame df = new DataPreparer.DataFrame(dateTimes, ints, strings); // This will throw if the columns are of different lengths
            Console.WriteLine(df);
            Console.WriteLine("Succeed!");
        }
    }
}
