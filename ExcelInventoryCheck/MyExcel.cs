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
        public Boolean isExcelFilePathValid = false;
        string path = string.Empty;
        string savingPath = string.Empty;
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

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws=wb.Worksheets[1];
        }

        public void CreateNewWorksheet()
        {
            wb.Worksheets.Add(After: ws);
        }

        /// <summary>
        /// Show Open Dialog box to select excel file and then open in in memeory
        /// </summary>
        public void OpenExcelFileWithDialog()
        {
            isExcelFilePathValid = false;
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Title = "select Excel file 2003";
            openDialog.Filter = "excel 97to2003(*.xls)|*.xls|excel2007 and above(*.xlsx)|*.xlsx";
            openDialog.FilterIndex = 1;

            //this use for #Test Only
            try
            {
                openDialog.InitialDirectory = @"E:\Project\Vahid";
            }
            catch (Exception)
            {
                
            }

            try
            {
                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    if (openDialog.FileName != null) this.path = openDialog.FileName;

                    OpenExcelFile(path, 1);
                    isExcelFilePathValid = true;

                }

            }
            catch (Exception)
            {

              
            }

        }
        
        public void SaveExcelFileWithDialog()
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Title = "select Excel file 2003";
            saveDialog.Filter = "excel 97to2003(*.xls)|*.xls|excel2007 and above(*.xlsx)|*.xlsx";
            saveDialog.FilterIndex = 1;
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                if (saveDialog.FileName != null) this.savingPath = saveDialog.FileName;
                this.SaveAs(this.savingPath);
                //OpenExcelFile(path, 1);

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
                
                var cellValue= ws.Cells[i, j].value2;

                //for test only
                //MessageBox.Show(cellValue.ToString());

                return cellValue.ToString();
            }
            return string.Empty;
        }


        public void Save()
        {
            wb.Save();

        }
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }
        /// <summary>
        /// write sring in excel file
        /// </summary>
        /// <param name="i"></param>
        /// <param name="j"></param>
        /// <param name="valueShouldToWrite"></param>
        public void WriteToCell(int i ,int j ,string valueShouldToWrite)
        {
            ws.Cells[i, j].value2 = valueShouldToWrite;

        }

        /// <summary>
        /// Crate Prouct From Excel Row
        /// </summary>
        /// <param name="row"></param>
        /// <returns>obj of Product</returns>
        public ExcelInventoryCheck.frmMain.Prouduct ReadProuct(int row)
        {
            ExcelInventoryCheck.frmMain.Prouduct myProduct=new ExcelInventoryCheck.frmMain.Prouduct() ;
            string productCode=ReadCell(row, 2);
            string productName = ReadCell(row, 1); 

            if (string.IsNullOrEmpty(productName))
            {
                //flag proudut to count 
                myProduct.code = "-1";
                return myProduct;
            }
            myProduct.code = productCode;
            myProduct.name = productName;

            //handel empty cell for column initial number
            string initValueString = ReadCell(row, 4);
            //MessageBox.Show(initValueString);
            if (initValueString!=string.Empty || initValueString != null)
            {
                try
                {
                    myProduct.initialNumber = int.Parse(initValueString);
                }
                catch (Exception)
                {

                    myProduct.initialNumber = 0;
                }
                
            }

            return myProduct;

        }



        public void WriteProuduct(ExcelInventoryCheck.frmMain.Prouduct myProduct,int rowIndexforFiling)
        {
            //int numberOfFilledRow = ws.Rows.Count;
            //int rowIndexforFiling = numberOfFilledRow++;
            WriteToCell(rowIndexforFiling+2, 1, myProduct.name);
            WriteToCell(rowIndexforFiling+2, 2, myProduct.code);
            WriteToCell(rowIndexforFiling+2, 3, myProduct.initialNumber.ToString());
            WriteToCell(rowIndexforFiling+2, 4, myProduct.totalNumber.ToString());
            WriteToCell(rowIndexforFiling+2, 5, myProduct.diffrenceNumber.ToString());

        }
        public void WriteHeder()
        {
            WriteToCell(1, 1, "نام کالا");
            WriteToCell(1, 2, "بارکد");
            WriteToCell(1, 3, "تعداد اولیه");
            WriteToCell(1, 4, "تعداد شمارش شده");
            WriteToCell(1, 5, "اختلاف");


        }


        public void WriteProducutList(List<ExcelInventoryCheck.frmMain.Prouduct> myProductList)
        {
            int numberOfProudct = 0;
            foreach (ExcelInventoryCheck.frmMain.Prouduct product in myProductList)
            {
                numberOfProudct++;
                WriteProuduct(product, numberOfProudct);
            }
        }

        public void WriteProducutList(Dictionary<string, ExcelInventoryCheck.frmMain.Prouduct> myProductList)
        {
            try
            {
                WriteHeder();
                for (int index = 0; index <= myProductList.Count; index++)
                {
                    var item = myProductList.ElementAt(index);
                    WriteProuduct(item.Value, index);
                }
            }
            catch (Exception)
            {

               
            }

        }

        /// <summary>
        /// az in tabe bara ijad ye list estefade mikonam
        /// </summary>
        /// <returns></returns>
        public List<ExcelInventoryCheck.frmMain.Prouduct> ReadExcelDataBase()
        {
            List<ExcelInventoryCheck.frmMain.Prouduct> myDataBaseList = new List<ExcelInventoryCheck.frmMain.Prouduct>();
            ExcelInventoryCheck.frmMain.Prouduct currentProuct = new ExcelInventoryCheck.frmMain.Prouduct();


            //#ForTest
            //MessageBox.Show(ws.Rows.Count.ToString());

            // startubg index
            try
            {
               // MessageBox.Show(ws.Rows.Count.ToString());
                for (int i = 2; i < ws.Rows.Count; i++)
                {
                    currentProuct = ReadProuct(i);
                    if (currentProuct.code == "-1")
                    {
                        break;
                    }
                    myDataBaseList.Add(currentProuct);
                }
                return myDataBaseList;
            }
            catch (Exception)
            {
                return myDataBaseList;

            }

         }


        public void Close()
        {
            wb.Close();
        }
    }
}
