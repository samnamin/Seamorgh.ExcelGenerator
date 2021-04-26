﻿using System.Collections.Generic;
using System.Data;
using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Row
    {
        public Row()
        {
            Columns = new List<Column>();
        }

        public DataRowCollection Data { get; set; }
        public Location StartLocation { get; set; }
        public Location EndLocation { get; set; }
        public Color BackColor { get; set; } = Color.White;
        public Color ForeColor { get; set; } = Color.Black;
        public List<Column> Columns { get; set; }
        public int Height { get; set; }
        public List<string> MergedCells { get; set; }
        public Border InLineBorder { get; set; }
        public Border OutLineBorder { get; set; }
        public bool IsBordered { get; set; }
        public string Formulas { get; set; }
        public int ColumnsCount
        {
            get
            {
                return Columns.Count;
            }
        }
        public Location NextHorizontalLocation
        {
            get
            {
                return new Location(EndLocation.X + 1, EndLocation.Y - StartLocation.Y);
            }
        }

        public Column AddColumn()
        {
            Column column = new(NextVerticalLocation);
            Columns.Add(column);
            return column;
        }

        public Location NextVerticalLocation
        {
            get
            {
                return new Location(EndLocation.X + StartLocation.X, EndLocation.Y + 1);
            }
        }

        public Column GetColumn(int X)
        {
            return Columns[X + 1];
        }
    }
}
