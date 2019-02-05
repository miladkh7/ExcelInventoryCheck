using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using ExcelInventoryCheck;
namespace ExcelInventoryCheck
{
    class MyExcel
    {

        string path = string.Empty;
        _Application excel = new _Excel.Application();

        Workbook wb;
        Worksheet ws;
        /// <summary>
        /// bara modiriat har file excel iz in class estefade mikonikm
        /// </summary>
        /// <param name="path">in parametr adrese file excel hastes</param>
        /// <param name="sheet">shomare sheet in file hastes</param>
        public MyExcel(string path, int sheet)
        {


            OpenExcelFile(path, sheet);

        }
        public MyExcel()
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="sheet"></param>
        public void OpenExcelFile(string path, int sheet)
        {
            this.path = path;
            this.wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        /// <summary>
        /// Show Open Dialog box to select excel file and then open in in memeory
        /// </summary>
        public void OpenExcelFileWithDialog()
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Title = "select Excel file 2003";
            openDialog.Filter = "excel 97to2003(*.xls)|*.xls|excel2007 and above(*.xlsx)|*.xlsx";
            openDialog.FilterIndex = 1;

            //this use for #Test Only
            openDialog.InitialDirectory = @"E:\Project\Vahid";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                if (openDialog.FileName != null) this.path = openDialog.FileName;

                OpenExcelFile(path, 1);

            }


        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="i">index of Row</param>
        /// <param name="j">index of column</param>
        /// <returns>value of Cell in row i and column j</returns>
        public string ReadCell(int i, int j)
        {
            
            //check index of Row and culumn
            if (i < 1 || j < 1)
            {
                MessageBox.Show("indices must be greater that zero");
                return string.Empty;
            }
            if (ws.Cells[i, j].value2 != null)
            {
                return ws.Cells[i, j].value2;
            }
            return string.Empty;
        }
       

    }
}
