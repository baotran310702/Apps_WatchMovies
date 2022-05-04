using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataDB = System.Data.OleDb;

namespace BTTH1
{
    public partial class formDetails : Form
    {
        public formDetails()
        {
            InitializeComponent();
        }
        
        private string thisPath = Directory.GetCurrentDirectory();

        static string[] filmData = new string[9];

        public formDetails(string[] s)
        {
            InitializeComponent();
            this.Icon = new Icon(thisPath + "\\logo\\Itube.ico");
            tenPhim.Text = s[1];
            TacGia.Text = s[2];
            QuocGia.Text = s[3];
            TheLoai.Text = s[4];
            LuotXem.Text = s[7];
            MoTa.Text = s[6];

            Image img = Image.FromFile(thisPath +"\\img\\"+ s[8] + ".png");
            pictureBox3.Image = img;
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            tenFile = s[8];
            idPhim = s[0];
            int n = Int32.Parse(s[5]);
            showStar(n);
            filmData = s;
        }
        //Khởi tạo biến để lưu tên file, idphim 
        static string tenFile = "";
        static string idPhim = "";  

        //Show the amount of film's star
        private void showStar(int n)
        {
            listView2.Items.Clear();
            for (int i = 1; i <= n; i++)
            {
                listView2.Items.Add("", 0);
            }
            for (int j = n + 1; j <= 5; j++)
            {
                listView2.Items.Add("", 1);
            }
            listView2.LargeImageList = imageList2;
            listView2.View = View.LargeIcon;
        }

        //This method is made to update star when clicking!
        private void UpdateStar(string id, int numStar)
        {
            string path = "Data.xlsx";
            var package = new ExcelPackage(new FileInfo(path));

            string[] s = new string[9];

            //Làm việc ở worksheets đầu 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            //Duyệt tuần tự từ dòng 2 tới dòng cuối 

            for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
            {
                try
                {
                    //Biến j thể hiện các cột
                    for (int j = 1; j < 10; j++)
                    {
                        string str = worksheet.Cells[i, j].Value.ToString();

                        if (str == idPhim)
                        {
                            DataDB.OleDbConnection MyCnt;
                            DataDB.OleDbCommand cmd = new DataDB.OleDbCommand();
                            String sql = null;
                            //string filePath = "C:\\Users\\DELL\\Desktop\\C#\\Đồ án\\Data.xlsx";
                            MyCnt = new DataDB.OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = '" + "Data.xlsx" + "';Extended Properties=\"Excel 12.0;HDR=YES;\"");
                            MyCnt.Open();
                            cmd.Connection = MyCnt;
                            sql = "update [sheet1$] set DanhGia = " + numStar.ToString() + " where MaPhim = '" + idPhim + "'";
                            cmd.CommandText = sql;
                            cmd.ExecuteNonQuery();
                            MyCnt.Close();
                        }
                    }
                }
                catch
                {

                }
            }
        }

        //This method for update view when watch_btn is clicked!
        private void TangView(string idPhim)
        {

            //string path = @"C:\Users\DELL\Desktop\C#\Đồ án\Data.xlsx";
            var package = new ExcelPackage(new FileInfo("Data.xlsx"));

            string[] s = new string[9];

            //Làm việc ở worksheets đầu 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            //Duyệt tuần tự từ dòng 2 tới dòng cuối 

            for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
            {
                try
                {
                    //Biến j thể hiện các cột
                    for (int j = 1; j < 10; j++)
                    {
                        string str = worksheet.Cells[i, j].Value.ToString();

                        if (str == idPhim)
                        {
                            string viewStr = worksheet.Cells[i, 8].Value.ToString();
                            int viewplus = Int32.Parse(viewStr);
                            viewplus++;
                            viewStr = viewplus.ToString();

                            //cExcel.Application excel = new cExcel.Application();
                            //cExcel.Workbook sheet = excel.Workbooks.Open(@"C:\Users\DELL\Desktop\C#\Đồ án\Data.xlsx");
                            //cExcel.Worksheet x = excel.ActiveSheet as cExcel.Worksheet;

                            //cExcel.Range userRange = x.UsedRange;
                            //x.Cells[i,8] = viewStr;
                            //sheet.Save();
                            //sheet.Close(true, Type.Missing, Type.Missing);
                            //excel.Quit();
                            DataDB.OleDbConnection MyCnt;
                            DataDB.OleDbCommand cmd = new DataDB.OleDbCommand();
                            String sql = null;
                            //string filePath = "C:\\Users\\DELL\\Desktop\\C#\\Đồ án\\Data.xlsx";
                            MyCnt = new DataDB.OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = '" + "Data.xlsx" + "';Extended Properties=\"Excel 12.0;HDR=YES;\"");
                            MyCnt.Open();
                            cmd.Connection = MyCnt;
                            sql = "update [sheet1$] set LuotView = " + viewplus.ToString() + " where MaPhim = '" + idPhim + "'";
                            cmd.CommandText = sql;
                            cmd.ExecuteNonQuery();
                            MyCnt.Close();
                        }
                    }
                }
                catch
                {

                }
            }

        }

        //Function btn

        private void listView2_MouseClick(object sender, MouseEventArgs e)
        {
            ListViewItem list = listView2.GetItemAt(e.X, e.Y);

            int index = list.Index + 1;

            UpdateStar(idPhim, index);

            showStar(index);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Form1 form = new Form1();
            form.Show();
            this.Hide();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Form1 form = new Form1();
            form.Show();
            this.Hide();
        }

        private void label2_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Form1 form = new Form1();
            form.Show();
            this.Hide();
        }

        private void btnWatch_Click(object sender, EventArgs e)
        {
            TangView(filmData[0]);
            addHistory();
            fmShow fm = new fmShow(filmData);
            Cursor = Cursors.WaitCursor;
            fm.Show();
            this.Hide();
        }

        //addHistory

        private void addHistory()
        {
            string path = "Data.xlsx";
            var package = new ExcelPackage(new FileInfo(path));

            //Làm việc ở worksheets đầu 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            try
            {
                DataDB.OleDbConnection MyCnt;
                DataDB.OleDbCommand cmd = new DataDB.OleDbCommand();
                String sql = null;
                //string filePath = "C:\\Users\\DELL\\Desktop\\C#\\Đồ án\\Data.xlsx";
                MyCnt = new DataDB.OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = '" + "Data.xlsx" + "';Extended Properties=\"Excel 12.0;HDR=YES;\"");
                MyCnt.Open();
                cmd.Connection = MyCnt;
                sql = "insert into [sheet2$](idPhim,namePhim,timeWatch,danhGia,soLan) values ('" + filmData[0].ToString() + "','" + filmData[1].ToString() + "','" + DateTime.Now.ToString() + "'," + filmData[5].ToString() + "," + filmData[7].ToString() + ")";
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
                MyCnt.Close();
            }
            catch
            {
                MessageBox.Show("err");
            }
        }

        
        private void btnBack_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Form1 form = new Form1();
            form.Show();
            this.Hide();
        }

        private void newPhimLabel_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Form1 form = new Form1();
            form.Show();
            this.Hide();
        }

        private void formDetails_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void formDetails_Load(object sender, EventArgs e)
        {

        }
        int pos = 0;
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                searchList.Visible = false;
            }
            else
            {
                searchList.Visible=true;
                searchList.Clear();
                pos = 0;
                string[] s = new string[9];

                searchList.Clear();

                var package = new ExcelPackage(new FileInfo("Data.xlsx"));

                //Làm việc ở worksheets đầu 
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                //Quét dữ liệu để lấy ra ảnh bìa và tên phim
                int i = worksheet.Dimension.Start.Row;
                while (i <= worksheet.Dimension.End.Row)
                {
                    s = getName(textBox1.Text);
                    try
                    {
                        if (s != null)
                        {
                            string[] maphim = s[0].Split('P');
                            int phim_key = Int32.Parse(maphim[1]);
                            phim_key--;
                            searchList.Items.Add(s[1], phim_key);
                            searchList.LargeImageList = imageList1;
                            searchList.View = View.LargeIcon;

                        }
                    }
                    catch { }

                    i++;
                }
            }
            
        }

        private string[] getName(string fName)
        {
            string[] s = new string[9];

            string path = "Data.xlsx";
            var package = new ExcelPackage(new FileInfo(path));

            //Làm việc ở worksheets đầu 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            //Duyệt tuần tự từ dòng 2 tới dòng cuối 

            for (int i = pos; i <= worksheet.Dimension.End.Row; i++)
            {
                try
                {
                    //Biến j thể hiện các cột
                    //Chỉ cần check cột tên phim thôi
                    string str = worksheet.Cells[i, 2].Value.ToString();
                    string str1 = str.ToLower();
                    string fName1 = fName.ToLower();
                    if (str1.Contains(fName1))
                    {
                        for (int j = 1; j < 10; j++)
                        {
                            string rowValue = worksheet.Cells[i, j].Value.ToString();
                            s[j - 1] = rowValue;
                        }
                        pos = i + 1;
                        return s;
                    }
                }
                catch
                {

                }
            }
            return null;
        }

        private void searchList_MouseClick(object sender, MouseEventArgs e)
        {
            ListViewItem list = searchList.GetItemAt(e.X, e.Y);
            //MessageBox.Show(list.ImageIndex.ToString());
            int index = list.ImageIndex + 1;
            string[] s = getData(index);

            formDetails form = new formDetails(s);
            form.Show();
            this.Hide();
        }

        private string[] getData(int index)
        {
            var package = new ExcelPackage(new FileInfo("Data.xlsx"));

            string[] s = new string[9];

            //Làm việc ở worksheets đầu 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            //Duyệt tuần tự từ dòng 2 tới dòng cuối 

            for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
            {
                try
                {
                    //Biến j thể hiện các cột

                    string str = worksheet.Cells[i, 1].Value.ToString();
                    string[] maphim = str.Split('P');
                    int phim_key = Int32.Parse(maphim[1]);
                    if (phim_key == index)
                    {
                        for (int j = 1; j < 10; j++)
                        {

                            s[j - 1] = worksheet.Cells[i, j].Value.ToString();

                        }
                    }
                }
                catch
                {

                }
            }
            return s;
        }
    }
}
