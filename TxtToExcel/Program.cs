using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace TxtToExcel
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            TxtToExcel("E:\\git\\01.开关机记录\\01.开关机数据\\开关机数据.log", "E:\\git\\TxtToExcel\\TxtToExcel\\bin\\Debug\\开关机数据.xls");
            //Application.Run(new Form1());
        }

        public static void TxtToExcel(string TxtFilePath, string OutputPath)
        {
            try
            {
                //指定文件的编码方式
                StreamReader sr = new StreamReader(TxtFilePath,Encoding.Default);
                string strLine = sr.ReadLine();
                //sr.Close();

                //将指定内容写入指定文件
                //StreamWriter sw = new StreamWriter(OutputPath);
                //sw.WriteLine(strLine);
                //sw.Close();

                int rowNum = 1;
                object missing = System.Reflection.Missing.Value;
                Excel.Application app = new Excel.ApplicationClass();
                app.Application.Workbooks.Add(true);
                //新建一个Workbook
                Excel.Workbook book = app.ActiveWorkbook;
                //新建Worksheet对象
                Excel.Worksheet sheet = (Excel.Worksheet)book.ActiveSheet;
                while (!string.IsNullOrEmpty(strLine))
                {
                    string[] tempArr;
                    //文本分割
                    tempArr = strLine.Split(new char[] { ':', ' ' }, 4, StringSplitOptions.RemoveEmptyEntries);
                    for (int k = 1; k <= tempArr.Length; k++)
                    {
                        sheet.Cells[rowNum, k] = tempArr[k - 1];
                    }
                    strLine = sr.ReadLine();
                    rowNum++;

                }
                Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(sheet.Cells[1, 4], sheet.Cells[rowNum, 4]);
                range.NumberFormat = "hh:mm:ss";//设置时间格式hh:mm:ss  
                //RowHeight   "1:1"表示第一行, "1:2"表示,第一行和第二行 
                ((Excel.Range)sheet.Rows["1:2", System.Type.Missing]).RowHeight = 30;
                //ColumnWidth "A:B"表示第一列和第二列, "A:A"表示第一列
                ((Excel.Range)sheet.Columns["A:D", System.Type.Missing]).ColumnWidth = 15;
                //设置字体大小
                Excel.Range excelRange = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[rowNum, 4]);
                excelRange.Font.Size = 15;
                //保存excel文件  
                book.SaveCopyAs(OutputPath);
                //关闭文件  
                //book.Close(false, missing, missing);
                //退出excel  
                // app.Quit();
                MessageBox.Show("转化成功！");
                //使Excel可视
                app.Visible = true;


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

   

    }
}
