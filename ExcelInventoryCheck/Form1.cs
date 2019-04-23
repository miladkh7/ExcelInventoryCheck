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
    public partial class frmMain : Form
    {
        #region class and structre
        /// <summary>
        /// in tarife har yek mahsol ast ke shame nam barcode va tedad kol mibashad
        /// </summary>
        public struct Prouduct
        {
            // declare name of product
           public string name;

            //declare barcode of prouct
           public string code;

            //declare number that we callculate number of them
           public int totalNumber ;

            //declare initila number of product that exist in first
           public int initialNumber;

           //declare defirence beetween init diffrcene and callculated number
           public int diffrenceNumber;

            /// <summary>
            /// 
            /// </summary>
            /// <returns>string contain name  and code</returns>
            public override string ToString()
            {
                return name.ToString() + " " + code.ToString();
            }
        }
        #endregion
        //declare dictinaroy of product Codes that we use them
        Dictionary<string, Prouduct> MyProducts = new Dictionary<string, Prouduct>();
        //declare list Recive Products 
        List<Prouduct> DataBase = new List<Prouduct>();

        //declare outPut list for review and save
        Dictionary<string, Prouduct> UniqListOfProduct;

        //declare refrence table
        MyExcel refrenceDataBase; 

        /// <summary>
        /// in tabe ye dictionary az file haye ke az list ke az file excel sakhtim tahvil ma mide
        /// </summary>
        /// <param name="productListFromExcel"></param>
        /// <returns></returns>
        public Dictionary<string, Prouduct> CraateDictionaryFromDataBaseList(List<Prouduct> productListFromExcel)
        {
            Dictionary<string, Prouduct> MyProducts = new Dictionary<string, Prouduct>();
            foreach (Prouduct item in productListFromExcel)
            {
                // add key if is not exist in the list
                if (!MyProducts.ContainsKey(item.code))
                {
                    Prouduct newProduct = item;
                    newProduct.totalNumber = 1;
                    newProduct.initialNumber = item.initialNumber;
                    newProduct.diffrenceNumber = item.initialNumber-1;
                    MyProducts.Add(item.code, newProduct);
                }
                //if exist we should callculate number of them
                else if(MyProducts.ContainsKey(item.code))
                {
                    var productSelectedToChangeInNumber = MyProducts[item.code];
                    productSelectedToChangeInNumber.totalNumber++;

                    //calculate number of proudct
                    productSelectedToChangeInNumber.diffrenceNumber =   productSelectedToChangeInNumber.initialNumber- productSelectedToChangeInNumber.totalNumber;

                    MyProducts[item.code] = productSelectedToChangeInNumber;
                    
                }

            }
            return MyProducts;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="datagrid"></param>
        /// <returns></returns>
        public List<Prouduct> CollectDataFromDataGrid(DataGridView datagrid)

        {
            List<Prouduct> listOfProuduct = new List<Prouduct>();
            Prouduct selectedToAddInList = new Prouduct();
            for (int i = 0; i < datagrid.Rows.Count-1; i++)
            {
                selectedToAddInList.code = datagrid.Rows[i].Cells[0].Value.ToString();
                selectedToAddInList.name= datagrid.Rows[i].Cells[1].Value.ToString();
                selectedToAddInList.initialNumber= int.Parse(datagrid.Rows[i].Cells[2].Value.ToString());
                listOfProuduct.Add(selectedToAddInList);
            }
            return listOfProuduct;
        }


        /// <summary>
        /// use dictionary of
        /// </summary>
        /// <param name="dataGridToFill"></param>
        /// <param name="uniqList"></param>
        public void EnterListInDataGridView(DataGridView dataGridToFill, Dictionary<string, Prouduct> uniqList)
        {

            foreach (var item in uniqList)
            {
                
                dataGridToFill.Rows.Add(item.Key, item.Value.name, item.Value.initialNumber);
            }
        }

        public void EnterListInDataGridViewFulDetail(DataGridView dataGridToFill, Dictionary<string, Prouduct> uniqDetailedList)
        {

            foreach (var item in uniqDetailedList)
            {

                dataGridToFill.Rows.Add(item.Key, item.Value.name,
                    item.Value.initialNumber,
                    item.Value.totalNumber,
                    item.Value.diffrenceNumber);
            }
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
            Prouduct myProduct = new Prouduct();
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


        /// <summary>
        /// Compair two list(dictionary of prouducts)
        /// </summary>
        /// <param name="listOne"> list mabda</param>
        /// <param name="listTwo">liste maghsad</param>
        /// <returns></returns>
        /// 
        private Dictionary<string, Prouduct> CompareList(Dictionary<string, Prouduct> listOne, Dictionary<string, Prouduct> listTwo)
        {
            Dictionary<string, Prouduct> compairResult = listOne;

            

            return compairResult;
            
        }
        private Dictionary<string, Prouduct> CompareList(Dictionary<string, Prouduct> listOne)
        {
            //foreach (var item in listOne)
            //{
               
            //}
            return listOne;

        }

        public frmMain()
        {
            
            InitializeComponent();
        }

        /// <summary>
        /// for saving reuslt of compair to an excel file
        /// </summary>
        /// <param name="ListToBeSave"></param>
        /// <param name="path"></param>

        private void SaveExcelFile(Dictionary<string, Prouduct> ListToBeSave ,string path)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            tslCurrentStatus.Text = "برای شروع فایلی را انتخاب کنید.";
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
        private void InterRowInDataGridView(DataGridView selectedDataGridView,string  ProuductCode,int row)
        {
            //MessageBox.Show(MyProducts[ProuductCode].code);
            try
            {
                selectedDataGridView[0, row].Value = MyProducts[ProuductCode].code;
                selectedDataGridView[1, row].Value = MyProducts[ProuductCode].name;
                selectedDataGridView[2, row].Value = MyProducts[ProuductCode].initialNumber;
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
            InterRowInDataGridView(dataGridView1, dataGridView1[0, rowIndex].Value.ToString(), rowIndex);
        }

        private void btnReciveDataBase_Click(object sender, EventArgs e)
        {
            tslCurrentStatus.Text="فایل مبدا(دیتا بیس رو انتخاب نمایید.";
            refrenceDataBase = new MyExcel();
            refrenceDataBase.OpenExcelFileWithDialog();
            tslCurrentStatus.Text = "در حال بارگذاری دیتابیس . لطفا صبور باشید";
            List<Prouduct> dataBase = refrenceDataBase.ReadExcelDataBase();
            MyProducts = CraateDictionaryFromDataBaseList(dataBase);
            tslCurrentStatus.Text = "بارگذاری تکمیل شد ،حالا می توانید نصبت با استعلام گیری اقدام نمایید.";
            dataGridView1.Focus();

            refrenceDataBase.Close();

        }


        private void btnCollectData_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Refresh();
            List<Prouduct> fullListOfProductEnterd = CollectDataFromDataGrid(dataGridView1);
            UniqListOfProduct = CraateDictionaryFromDataBaseList(fullListOfProductEnterd);

            //#TODO: compare between of them
            EnterListInDataGridViewFulDetail(dataGridView2, UniqListOfProduct);

        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        private void frmMain_SizeChanged(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnClearList_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
        }

        private void btnSaveResult_Click(object sender, EventArgs e)
        {
            tslCurrentStatus.Text = "محل ذخیره را انتخاب نمایید.";
            MyExcel outputFile = new MyExcel();
            outputFile.CreateNewFile();
            outputFile.WriteProducutList(UniqListOfProduct);
            tslCurrentStatus.Text = "در حال ذخیره نتایج";
            outputFile.SaveExcelFileWithDialog();
            tslCurrentStatus.Text = "نتایج با موفقیت ذخیره گردید";
            outputFile.Close();

        }


    }
 
    
 
}
