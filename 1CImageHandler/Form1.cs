using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office;
using System.Xml;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.IO;

namespace _1CImageHandler
{
    public partial class Form1 : Form
    {
        bool AutoStart = false;
        string fileNameIn;
        string fileNameOut;
        string patch;

        public Form1(string p1, string p2)
        {
            InitializeComponent();
            if (!String.IsNullOrWhiteSpace(p1) && !String.IsNullOrWhiteSpace(p2))
            {
                fileNameIn = p1;
                fileNameOut = p2;
                AutoStart = true;
            }
            else 
            {
                fileNameIn = Properties.Settings.Default.patchIn;
                fileNameOut = Properties.Settings.Default.patchOut;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dgvResult.Visible = false;
            txtFileIn.Text = fileNameIn;
            txtFileOut.Text = fileNameOut;
            patch = AppDomain.CurrentDomain.BaseDirectory;
            if (AutoStart) Process_Start(true);
        }

        private static DataTable ExcelToDataTable(string patch)
        {
            DataTable DTResult = new DataTable();
            if (!String.IsNullOrWhiteSpace(patch))
            {
                Microsoft.Office.Interop.Excel.Application objXL = null;
                Microsoft.Office.Interop.Excel.Workbook objWB = null;
                try
                {
                    objXL = new Microsoft.Office.Interop.Excel.Application();
                    objWB = objXL.Workbooks.Open(patch);
                    foreach (Microsoft.Office.Interop.Excel.Worksheet objSHT in objWB.Worksheets)
                    {
                        int rows = objSHT.UsedRange.Rows.Count;
                        int cols = objSHT.UsedRange.Columns.Count;
                        int noofrow = 1;
                        for (int c = 1; c <= cols + 1; c++)
                        {
                            string colname = objSHT.Cells[1, c].Text;
                            DTResult.Columns.Add(colname);
                            noofrow = 2;
                        }
                        for (int r = noofrow; r <= rows; r++)
                        {
                            DataRow dr = DTResult.NewRow();
                            for (int c = 1; c <= cols + 1; c++)
                            {
                                dr[c - 1] = objSHT.Cells[r, c].Text;
                            }
                            DTResult.Rows.Add(dr);
                        }
                    }
                    objWB.Close();
                    objXL.Quit();
                }
                catch
                {
                    objWB.Saved = true;
                    objWB.Close();
                    objXL.Quit();
                }
            }
            return DTResult;
        }

        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        private void btnXlsLoad_Click(object sender, EventArgs e)
        {
            dgvResult.DataSource = ExcelToDataTable(fileNameIn);

            DataGridViewImageColumn img = new DataGridViewImageColumn();
            dgvResult.Columns.Insert(11,img);
            img.HeaderText = "Image";
            img.Name = "img";
            foreach (DataGridViewRow row in dgvResult.Rows)
                row.Cells[11].Value = new Bitmap(1, 1);
        }

        private void Process_Start(bool close) 
        {
            if (System.IO.File.Exists(fileNameOut))
                System.IO.File.Delete(fileNameOut);
               
            if (System.IO.File.Exists(fileNameIn))
                System.IO.File.Copy(fileNameIn, fileNameOut);

            var xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileNameOut);
            Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[1];

            xlWorkSheet.Unprotect();

            int i_max = 0;
            int i_sum = 0;
            int i_start_pos = 0;
            int i_stop_pos = 0;
            int i_count = 0;

            Excel.Range usedRange = xlWorkSheet.UsedRange;
            foreach (Excel.Range row in usedRange.Rows)
                i_max++;

            //Зададим ширину столбца путь к картинке
            var SomeCell1 = (Excel.Range)xlWorkSheet.Cells[1, 11];
            SomeCell1.ColumnWidth = 0;

            //Зададим ширину столбца картинок
            var SomeCell2 = (Excel.Range)xlWorkSheet.Cells[1, 12];
            SomeCell2.ColumnWidth = 15;

            //Обработка позиций
            for (int i = 1; i <= i_max; i++)
            {
                var cellValue = (string)(xlWorkSheet.Cells[i, 11] as Excel.Range).Value;
                if (!String.IsNullOrWhiteSpace(cellValue))
                {
                    i_count++;
                    if (i_start_pos == 0) i_start_pos = i;

                    //Задаем высоту строки
                    var SomeCell = (Excel.Range)xlWorkSheet.Cells[i, 11];
                    SomeCell.RowHeight = 40;

                    //Вставляем картинку
                    if (System.IO.File.Exists(cellValue))
                    {
                        Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[i, 12];
                        float Left = (float)((double)oRange.Left + 1);
                        float Top = (float)((double)oRange.Top + 1);

                        float Height = (float)((double)oRange.Height - 2);
                        float Witch = (float)((double)oRange.Width - 2);

                        xlWorkSheet.Shapes.AddPicture(cellValue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, Witch, Height);
                    }
                    //Вставляем формулу 
                    var SomeCellF = (Excel.Range)xlWorkSheet.Cells[i, 19];
                    SomeCellF.FormulaR1C1 = String.Format("=R{0}C5*R{0}C18", i);
                    i_stop_pos = i;
                    i_sum = i + 1;
                }
            }

            //Вставка формулы итого ко-во
            var SomeCellSumCount = (Excel.Range)xlWorkSheet.Cells[i_sum, 18];
            SomeCellSumCount.FormulaR1C1 = String.Format("=SUM(R[-{0}]C:R[-1]C)", i_count);

            //Вставка формулы итого сумма
            var SomeCellSumPrice = (Excel.Range)xlWorkSheet.Cells[i_sum, 19];
            SomeCellSumPrice.FormulaR1C1 = String.Format("=SUM(R[-{0}]C:R[-1]C)", i_count);

            //Вставка формулы вес
            var SomeCellSumVes = (Excel.Range)xlWorkSheet.Cells[i_sum, 19];
            SomeCellSumVes.FormulaR1C1 = String.Format("=SUM(R[-{0}]C:R[-1]C)", i_count);

            //Заблокировать на редактирование//R6C18
            xlWorkSheet.Range[String.Format("R{0}", i_start_pos), String.Format("R{0}", i_stop_pos)].Locked = false;//Выбранный разрешенный диапазон
            xlWorkSheet.Protect(UserInterfaceOnly: true);
            xlWorkBook.SaveAs(fileNameOut, Excel.XlFileFormat.xlWorkbookNormal);

            xlWorkBook.Close(true);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);

            if (close) this.Close();
        }
        private void btnStart_Click(object sender, EventArgs e)
        {
            Process_Start(false);
            Process.Start(patch);
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd1 = new OpenFileDialog();
            ofd1.DefaultExt = "*.xls;*.xlsx";
            ofd1.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd1.InitialDirectory = Application.StartupPath + "\\";
            ofd1.RestoreDirectory = true;

            if (ofd1.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали документ!");
            }
            else 
            {
                Properties.Settings.Default.patchIn= ofd1.FileName;
                Properties.Settings.Default.Save();

                txtFileIn.Text = Properties.Settings.Default.patchIn;

                Properties.Settings.Default.patchOut = ofd1.FileName.Replace(".xls", "_out.xls");
                Properties.Settings.Default.Save();

                txtFileOut.Text = Properties.Settings.Default.patchOut;

                dgvResult.DataSource = ExcelToDataTable(fileNameIn);

                DataGridViewImageColumn img = new DataGridViewImageColumn();

                dgvResult.Columns.Insert(11, img);
                img.HeaderText = "Image";
                img.Name = "img";

                foreach (DataGridViewRow row in dgvResult.Rows)
                   row.Cells[11].Value = new Bitmap(1, 1);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) 
            {
                dgvResult.Visible = true;
            }
        }
    }
}
