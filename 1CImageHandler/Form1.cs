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

        private void Process_Start(bool close) 
        {
            if (System.IO.File.Exists(fileNameOut))
                System.IO.File.Delete(fileNameOut);
               
            if (System.IO.File.Exists(fileNameIn))
                System.IO.File.Copy(fileNameIn, fileNameOut);

            var xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileNameOut);
            xlWorkBook.CheckCompatibility = false;

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

            //Объединим ячейки
            xlWorkSheet.get_Range("A1:B2").Merge();

            //Задаем высоту первой строки
            var SomeCellFirstRow = (Excel.Range)xlWorkSheet.Cells[1, 1];
            SomeCellFirstRow.RowHeight = 40;

            //Зададим ширину столбца путь к картинке
            var SomeCell1 = (Excel.Range)xlWorkSheet.Cells[1, 12];
            SomeCell1.ColumnWidth = 0;

            //Зададим ширину столбца картинок
            var SomeCell2 = (Excel.Range)xlWorkSheet.Cells[1, 13];
            SomeCell2.ColumnWidth = 15;

            //Зададим ширину столбца картинок
            var SomeCell3 = (Excel.Range)xlWorkSheet.Cells[1, 3];
            SomeCell3.ColumnWidth = 0;

            //Обработка позиций
            for (int i = 1; i <= i_max; i++)
            {
                if(xlWorkSheet.Cells[i, 3].Value!=null)
                {
                    var cellValue2 = (decimal)(xlWorkSheet.Cells[i, 3] as Excel.Range).Value;

                    if (cellValue2 == 101) 
                    {
                        i_count++;
                        if (i_start_pos == 0) i_start_pos = i;

                        var cellValue = (string)(xlWorkSheet.Cells[i, 12] as Excel.Range).Value;

                        //Задаем высоту строки
                        var SomeCell = (Excel.Range)xlWorkSheet.Cells[i, 12];
                        SomeCell.RowHeight = 50;

                        if (!String.IsNullOrWhiteSpace(cellValue))
                        {
                            //Вставляем картинку
                            if (System.IO.File.Exists(cellValue))
                            {
                                Debug.WriteLine("image patch="+ cellValue);
                                Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[i, 13];
                                float Left = (float)((double)oRange.Left + 1);
                                float Top = (float)((double)oRange.Top + 1);

                                float Height = (float)((double)oRange.Height - 2);
                                float Witch = (float)((double)oRange.Width - 2);
                                //" d:\\_1C_BD\\Торговля SQL СПБ\\Foto\\1.jpg"
                                xlWorkSheet.Shapes.AddPicture(cellValue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, Witch, Height);
                            }
                        }
                        //Вставляем формулу 
                        var SomeCellF = (Excel.Range)xlWorkSheet.Cells[i, 20];
                        SomeCellF.FormulaR1C1 = String.Format("=R{0}C6*R{0}C19", i);
                        i_stop_pos = i;
                        i_sum = i + 1;

                        //Вставляем формулу  =RC[-1]/RC[-6]*RC[3]
                        var SomeCellB = (Excel.Range)xlWorkSheet.Cells[i, 16];
                        SomeCellB.FormulaR1C1 = String.Format("=RC[-1]/RC[-8]*RC[3]", i);

                        //Вставляем формулу  =RC[-1]/RC[-8]*RC[1]
                        var SomeCellS = (Excel.Range)xlWorkSheet.Cells[i, 18];
                        SomeCellS.FormulaR1C1 = String.Format("=RC[-1]/RC[-10]*RC[1]", i);

                        i_stop_pos = i;
                        i_sum = i + 1;
                    }
                }
            }

            //Вставка формулы итого сумма
            var SomeCellSumPrice = (Excel.Range)xlWorkSheet.Cells[i_sum, 19];
            SomeCellSumPrice.FormulaR1C1 = String.Format("=SUM(R[-{0}]C[1]:R[-1]C[1])", i_count);

            //Вставка формулы итого Вес брутто =СУММ(R[-6]C[-4]:R[-2]C[-4])
            var SomeCellSumBrutto = (Excel.Range)xlWorkSheet.Cells[i_sum+1, 20];
            SomeCellSumBrutto.FormulaR1C1 = String.Format("=SUM(R[-{0}]C[-4]:R[-2]C[-4])", i_count+1);

            //Вставка формулы итого СВМ общая =СУММ(R[-7]C[-3]:R[-3]C[-3])
            var SomeCellSumSVM = (Excel.Range)xlWorkSheet.Cells[i_sum + 2, 20];
            SomeCellSumSVM.FormulaR1C1 = String.Format("=SUM(R[-{0}]C[-3]:R[-3]C[-3])", i_count+2);

            ////Объединим ячейки
            xlWorkSheet.get_Range(String.Format("A{0}:B{1}", i_max-1, i_max)).Merge();

            //Задаем высоту предпоследней строки строки
            var SomeCellEndRow = (Excel.Range)xlWorkSheet.Cells[i_max-1, 1];
            SomeCellEndRow.RowHeight = 40;

            //Заблокировать на редактирование//R6C18
            xlWorkSheet.Range[String.Format("S{0}", i_start_pos), String.Format("S{0}", i_stop_pos)].Locked = false;//Выбранный разрешенный диапазон
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

            }
        }
    }
}
