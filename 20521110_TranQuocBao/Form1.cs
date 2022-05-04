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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Icon = new Icon(thisPath +"\\logo\\Itube.ico");
            var package = new ExcelPackage(new FileInfo("Data.xlsx"));
            
            string[] s = new string[9];

            //Làm việc ở worksheets đầu 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            //Quét dữ liệu để lấy ra ảnh bìa và tên phim
            for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
            {
                s = getData(i);
                int index = i - 1;
                listView1.Items.Add(s[1], index);
            }

            listView1.LargeImageList = imageList1;
            listView1.View = View.LargeIcon;
        }

        string thisPath = Directory.GetCurrentDirectory();
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }      

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            ListViewItem list = listView1.GetItemAt(e.X, e.Y);
            //MessageBox.Show(list.ImageIndex.ToString());
            int index = list.ImageIndex + 1;
            string[] s = getData(index);

            formDetails form = new formDetails(s);
            form.Show();
            this.Hide();
        }

        //Hàm Lấy dữ liệu từ excel
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

        static int pos = 0;

        //getData From Name
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
                        for(int j=1;j<10;j++)
                        {
                            string rowValue = worksheet.Cells[i, j].Value.ToString();
                            s[j - 1] = rowValue;
                        }
                        pos = i+1;
                        return s;
                    }
                }
                catch
                {
                    
                }
            }
            return null;
        }


        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            pos = 0;
            string[] s = new string[9];

            listView1.Clear();

            var package = new ExcelPackage(new FileInfo("Data.xlsx"));

            //Làm việc ở worksheets đầu 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            //Quét dữ liệu để lấy ra ảnh bìa và tên phim
            int i = worksheet.Dimension.Start.Row;
            while ( i <= worksheet.Dimension.End.Row)
            {
                s = getName(textBox1.Text);
                try
                {
                    if (s!= null)
                    {
                        string[] maphim = s[0].Split('P');
                        int phim_key = Int32.Parse(maphim[1]);
                        phim_key--;
                        listView1.Items.Add(s[1], phim_key);
                        listView1.LargeImageList = imageList1;
                        listView1.View = View.LargeIcon;
                                                
                    }
                }
                catch { }
               
                i++;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            formHistory fHistory = new formHistory();
            fHistory.Show();
            this.Hide();
        }
    }
}
