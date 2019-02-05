using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelInventoryCheck
{
    public partial class Form1 : Form
    {
        #region class and structre
        /// <summary>
        /// in tarife har yek mahsol ast ke shame nam barcode va tedad kol mibashad
        /// </summary>
        public struct Prouduct
        {
           public string name;
           public string code;
           public int totalNumber ;
        }
        #endregion
        //declare dictinaroy of product Codes
        Dictionary<string, Prouduct> MyProducts = new Dictionary<string, Prouduct>();
        //declare list Recive Products 
        List<Prouduct> DataBase = new List<Prouduct>(); 


        public Dictionary<string, Prouduct> CraateDictionaryFromDataBaseList(List<Prouduct> productListFromExcel)
        {
            Dictionary<string, Prouduct> MyProducts = new Dictionary<string, Prouduct>();
            foreach (Prouduct item in productListFromExcel)
            {
                MyProducts.Add(item.code, item);
            }
            return MyProducts;
        }


        /// <summary>
        /// bara bedast ovordane time hejri shamsi
        /// </summary>
        /// <returns>tarih hejri shamsi</returns>
        public string GetCurrntHijriDate()
        {
            System.Globalization.PersianCalendar pc = new System.Globalization.PersianCalendar();
            string Date = pc.GetYear(DateTime.Now) + "/" + pc.GetMonth(DateTime.Now) + "/" + pc.GetDayOfMonth(DateTime.Now);
            string Time = DateTime.Now.ToString("hh:mm:ss");
            return Date.ToString()+" - "+ Time;

        }

        /// <summary>
        /// in tabe baraye ijad mahsole jadid mibashad
        /// </summary>
        /// <param name="productCode">barcode mahsol hastes</param>
        /// <param name="productName">esme mahsol hastes</param>
        /// <returns>ye shey ke meghadr mahsol ro neshon mide</returns>
        public Prouduct CreateProduct(string productCode,string productName)
        {
            Prouduct myProduct;
            myProduct.name = productName;
            myProduct.code = productCode;
            myProduct.totalNumber = 0;
            AddProductsToList(myProduct);
            return myProduct;
        }

        // this is only for Test
        public void CreateSomeRandomProduct()
        {
            CreateProduct("6260374016210", "cleanSing Gel");
            CreateProduct("6260682508797", "Creastal");
            CreateProduct("9000100829717", "Gliss");
            CreateProduct("6260955804557", "Loriant");
        }
        public void AddProductsToList(Prouduct selectedProduct)
        {
            MyProducts.Add(selectedProduct.code, selectedProduct);
        }

        public Form1()
        {
            
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //test load first excel file

            //string myFistExcelFile = @"e:\project\vahid\testexcel2.xls";
            //MyExcel excel1 = new MyExcel(myFistExcelFile, 1);
            //MessageBox.Show(excel1.ReadCell(0, 0).ToString());
            //#todo:check file is exist or not

            //OpenExcelFile();
            //MyExcel excel1 = new MyExcel();
            //excel1.OpenExcelFileWithDialog();
            //MessageBox.Show(excel1.ReadCell(1, 1).ToString());
            CreateSomeRandomProduct();
        }



        private void dataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
           
            int rowIndex = e.RowIndex;
            //MessageBox.Show( dataGridView1.Rows[rowIndex].Cells[0].Value.ToString());
                                    //MessageBox.Show(dataGridView1[1, 1].Value.ToString());
                                    //InterDataInDataGridView(dataGridView1, dataGridView1[0, rowIndex].Value.ToString(), rowIndex);


        }

        /// <summary>
        /// in method baraye vard kardne etelata mibashad
        /// </summary>
        /// <param name="selectedDataGridView"></param>
        /// <param name="ProuductCode">shamele code mahsol va sayer moshakhaste filed ha</param>
        /// <param name="row">shomare indexi ke bayad vard shavad</param>       
        private void InterDataInDataGridView(DataGridView selectedDataGridView,string  ProuductCode,int row)
        {
            //MessageBox.Show(MyProducts[ProuductCode].code);
            try
            {
                selectedDataGridView[0, row].Value = MyProducts[ProuductCode].code;
                selectedDataGridView[1, row].Value = MyProducts[ProuductCode].name;
                selectedDataGridView[2, row].Value = MyProducts[ProuductCode].totalNumber;
                selectedDataGridView[3, row].Value = GetCurrntHijriDate();
            }
            catch
            {
                selectedDataGridView[0, row].Value = ProuductCode;
                selectedDataGridView[1, row].Value ="Undefind Code";
                selectedDataGridView[2, row].Value = "0";
                selectedDataGridView[3, row].Value = GetCurrntHijriDate();
            }
            
        }


        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            //use this for show the value
            //MessageBox.Show(dataGridView1.Rows[rowIndex].Cells[0].Value.ToString());
            InterDataInDataGridView(dataGridView1, dataGridView1[0, rowIndex].Value.ToString(), rowIndex);
        }

        private void btnReciveDataBase_Click(object sender, EventArgs e)
        {

        }
    }
 
    
 
}
