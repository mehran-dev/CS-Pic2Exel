using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace pic2excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            FileDialog fd = new OpenFileDialog();

            fd.ShowDialog();


            if (!string.IsNullOrEmpty(fd.FileName))
            {
                //should be checked if extension is valid or image is valid).)
                pic1.Image = new Bitmap(fd.FileName);
            }


        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (pic1.Image == null)
            {
                MessageBox.Show("ابتدا عکسی را انتخاب کنید ");
                return;
            }



            var r = await ProcessFile();


        }


        public int CreateExcel()
        {
            int p,All = 1;
            int ctr=0;
            string ctrT;
            int width, height;
            Color pixelColor;
            FileInfo excelFile = new FileInfo(@"D:\test.xlsx");


            using (ExcelPackage excel = new ExcelPackage(excelFile))
            {
                var worksheet = excel.Workbook.Worksheets[1]; 
                    //excel.Workbook.Worksheets.Add("Worksheet1");
                //excel.Workbook.Worksheets.Add("Worksheet2");
                //excel.Workbook.Worksheets.Add("Worksheet3");


                // Create a Bitmap object from an image file.
                Bitmap myBitmap = new Bitmap(pic1.Image);

                // Get the color of a pixel within myBitmap.

                width = myBitmap.Width;
                height = myBitmap.Height;
                All = height * width;
              //  pBar1.Maximum = All;
                pBar1.Invoke(new MethodInvoker(delegate { pBar1.Maximum = All; }));
                MessageBox.Show(All.ToString());
                worksheet.Cells["A1:XFD1048576"].Style.Fill.PatternType = ExcelFillStyle.Solid;

                for (int w = 0; w < myBitmap.Width; w+=1)
                {
                    for (int h = 0; h < myBitmap.Height; h+=1)
                    {
                         pixelColor = myBitmap.GetPixel(w, h);

                        //set bg color pattern in xlsx
                        //worksheet.Cells[h+1,w+1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[h+1,w+1].Style.Fill.BackgroundColor.SetColor(pixelColor);

                        //end  bg color pattern in xlsx
                        //excel.Save();
                        ctr += 1;
                      //  pBar1.Value = ctr;
                        pBar1.Invoke(new MethodInvoker(delegate { pBar1.Value = ctr; }));

                        //p = (int)ctr / All; 

                        //ctrT = "w:" + width + "     h: " + height + "    h*w :" + (height * width) + "     number:  " + ctr +" Percent "  + p.ToString() +"%";


                        //lbl1.Invoke(new MethodInvoker(delegate { lbl1.Text = ctrT; }));

                    }
                    excel.Save();
                }








                //  excel.SaveAs(excelFile);
            }
           // Thread.Sleep(6500);
            MessageBox.Show("Done!");
            return 1;
        }

        async Task<int> ProcessFile()
        {
            Task<int> task = new Task<int>(CreateExcel);
            task.Start();
            return await task;
        }


    }
}
