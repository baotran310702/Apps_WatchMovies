using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using DataDB = System.Data.OleDb;
using System.IO;

namespace BTTH1
{
    public partial class formHistory : Form
    {
        public class Film
        {
            public string idPhim { get; set; }
            public string namePhim { get; set; }
            public string timeWatch { get; set; }
            public string danhGia { get; set; }
            public string soLan { get; set; }
        }
        public formHistory()
        {
            InitializeComponent();
        }

        private void formHistory_Load(object sender, EventArgs e)
        {
            var package = new ExcelPackage(new FileInfo("Data.xlsx"));

            string[] s = new string[5];
            //Làm việc ở worksheets đầu 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            //Quét dữ liệu để lấy ra ảnh bìa và tên phim
            
            dgvHistory.DataSource = getData();
        }

        private List<Film> getData()
        {
            List<Film> film_list = new List<Film>();
            var package = new ExcelPackage(new FileInfo("Data.xlsx"));

            string[] s = new string[5];

            //Làm việc ở worksheets đầu 
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

            //Duyệt tuần tự từ dòng 2 tới dòng cuối 

            for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
            {
                try
                {
                    //Biến j thể hiện các cột
                    for (int j = 1; j < 6; j++)
                    {

                        s[j - 1] = worksheet.Cells[i, j].Value.ToString();

                    }
                    Film film = new Film();
                    film.idPhim = s[0];
                    film.namePhim = s[1];
                    film.timeWatch = DateTime.Now.ToString();
                    film.soLan = s[4];
                    film.danhGia = s[3];
                    film_list.Add(film);
                }
                catch
                {

                }
            }
            return film_list;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Hide();
        }

        private void formHistory_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
