using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelWriteAndMerge
{
    public class MergeCells : CodeActivity
    {

        [Category("Input")]
        [DisplayName("ExcelFilePath")]
        [Description("Enter excel file path")]
        [RequiredArgument]
        public InArgument<string> ExcelFilePath { get; set; }

        [Category("Input")]
        [DisplayName("SheetName")]
        [Description("Enter sheet name")]
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }

        [Category("Input")]
        [DisplayName("CellsRange")]
        [Description("Enter cell range to merge")]
        [RequiredArgument]
        public InArgument<string> RangeToMerge { get; set; }

        [Category("Output")]
        public OutArgument<bool> Status { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                var FilePath = ExcelFilePath.Get(context);
                var sheetToWork = SheetName.Get(context);
                var InputRange = RangeToMerge.Get(context);
                Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = xl.Workbooks.Open(FilePath);
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[sheetToWork];
                xl.Visible = false;
                Microsoft.Office.Interop.Excel.Range range = ws.Range[InputRange];
                range.Merge();
                wb.Save();
                wb.Close();
                Status.Set(context, true);
            }
            catch (Exception)
            {
                Status.Set(context, false);
            }
        }

        public class WriteCell : CodeActivity
        {
            [Category("Input")]
            [DisplayName("ExcelFilePath")]
            [Description("Enter excel file path")]
            [RequiredArgument]
            public InArgument<string> ExcelFilePath { get; set; }

            [Category("Input")]
            [DisplayName("SheetName")]
            [Description("Enter sheet name")]
            [RequiredArgument]
            public InArgument<string> SheetName { get; set; }

            [Category("Input")]
            [DisplayName("RowIndex")]
            [Description("Enter row index")]
            [RequiredArgument]
            public InArgument<int> RowPosition { get; set; }

            [Category("Input")]
            [DisplayName("ColumnIndex")]
            [Description("Enter column index")]
            [RequiredArgument]
            public InArgument<int> ColumnPosition { get; set; }

            [Category("Input")]
            [DisplayName("CellValue")]
            [Description("Enter value to update")]
            [RequiredArgument]
            public InArgument<string> cellValue { get; set; }

            [Category("Output")]
            public OutArgument<bool> Status { get; set; }

            protected override void Execute(CodeActivityContext context)
            {
                try
                {
                    var FilePath = ExcelFilePath.Get(context);
                    var sheetToWork = SheetName.Get(context);
                    var rowToUpdate = RowPosition.Get(context);
                    var columnToUpdate = ColumnPosition.Get(context);
                    var valueToUpdate = cellValue.Get(context).ToString();
                    Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook wb = xl.Workbooks.Open(FilePath);
                    Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[sheetToWork];
                    xl.Visible = false;
                    ws.Cells[rowToUpdate, columnToUpdate] = valueToUpdate;
                    wb.Save();
                    wb.Close();
                    Status.Set(context, true);
                }
                catch (Exception)
                {
                    Status.Set(context, false);
                }
            }
        }
    }
}
