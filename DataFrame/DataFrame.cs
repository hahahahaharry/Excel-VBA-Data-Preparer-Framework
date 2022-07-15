using Microsoft.Data.Analysis;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using DF = Microsoft.Data.Analysis.DataFrame;

namespace DataPreparer
{
    public class DataFrame
    {
        private DF dataframe;
        public DataFrame(params DataFrameColumn[] columns)
        {
            dataframe = new DF(columns);
        }

        public DataFrame LoadFromRange(Range range, bool hasHeader)
        {
            List<DataFrameColumn> dfCols = new List<DataFrameColumn>();
            StringDataFrameColumn dfColumn = null;
            foreach (Range col in range.Columns)
            {
                dfColumn = new StringDataFrameColumn(hasHeader?col.Cells[1][1]:"Column1", hasHeader?col.Count-1:col.Count);
                foreach(Range cell in col.Cells)
                {
                    dfColumn.Append((String)cell.Value2);
                }
                dfCols.Add(dfColumn);
            }
            return new DataFrame(dfCols.ToArray());
        }

        public override string ToString() 
        {
            return "override toString()";
        }
    }
}
