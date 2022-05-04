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

namespace BTTH1
{
    public partial class fmShow : Form
    {
        public fmShow()
        {
            InitializeComponent();
;       }

        private string thisPath = Directory.GetCurrentDirectory();

        public fmShow(string[] film)
        {
            InitializeComponent();
            this.Icon = new Icon(thisPath + "\\logo\\Itube.ico");
            axWindowsMediaPlayer1.URL = thisPath+ "\\video\\" + film[8] + ".mp4";
            axWindowsMediaPlayer1.Ctlcontrols.play();
            filmData = film;

        }
        
        static string[] filmData = new string[9];

        private string[] getData(string id)
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
                    for (int j = 1; j < 10; j++)
                    {
                        string str = worksheet.Cells[i, j].Value.ToString();
                        if (str==id)
                        {
                            for (j = 1; j < 10; j++)
                            {

                                s[j - 1] = worksheet.Cells[i, j].Value.ToString();

                            }
                        }

                    }
                }
                catch
                {

                }
            }
            return s;
        }

        private void back_btn_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            string[] sData = getData(filmData[0]);
            Form1 form = new Form1();
            form.Show();            
            axWindowsMediaPlayer1.Ctlcontrols.stop();        
            this.Hide();

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Hide();
            Application.Exit();
        }

        private void fmShow_Click(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
