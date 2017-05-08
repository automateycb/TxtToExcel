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
                sheet.Cells[rowNum, 1] = "日期";
                sheet.Cells[rowNum, 2] = "星期";
                sheet.Cells[rowNum, 3] = "开机时间";
                sheet.Cells[rowNum, 4] = "关机时间";
                sheet.Cells[rowNum, 5] = "运行时长";
                while (!string.IsNullOrEmpty(strLine))
                {
                    string[] tempArr;
                    //文本分割
                    tempArr = strLine.Split(new char[] { ':', ' ' }, 4, StringSplitOptions.RemoveEmptyEntries);
                    if (rowNum == 0)  //不成立   测试使用
                    {
                        sheet.Cells[rowNum, 1] = tempArr[1];  //日期
                        sheet.Cells[rowNum, 2] = tempArr[2];  //星期
                        if (tempArr[0] == "关机时间")
                        {
                            sheet.Cells[rowNum, 4] = tempArr[3];
                        }
                        else if (tempArr[0] == "开机时间")
                        {
                            sheet.Cells[rowNum, 3] = tempArr[3];
                        }
                        rowNum++;
                    }
                    else 
                    {
                        string date = ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[rowNum, 2]).Text.ToString();
                        string cells3 = ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[rowNum, 3]).Text.ToString();
                        string cells4 = ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[rowNum, 4]).Text.ToString();
                        //判断日期是否与前一行数据一样
                        if (tempArr[2] ==(date))    //一样 
                        {
                
                            if (cells3==null && cells4==null)  //前一行开关机时间都有
                            {
                                rowNum++;        //新起一行数据
                                sheet.Cells[rowNum, 1] = tempArr[1];  //日期
                                sheet.Cells[rowNum, 2] = tempArr[2];  //星期
                                if (tempArr[0] == "关机时间")
                                {
                                    sheet.Cells[rowNum, 4] = tempArr[3];
                                }
                                else if (tempArr[0] == "开机时间")
                                {
                                    sheet.Cells[rowNum, 3] = tempArr[3];
                                }
                            }
                            else    //开关机时间不是都有
                            {
                                sheet.Cells[rowNum, 1] = tempArr[1];  //日期
                                sheet.Cells[rowNum, 2] = tempArr[2];  //星期
                                if (tempArr[0] == "关机时间")
                                {
                                    sheet.Cells[rowNum, 4] = tempArr[3];
                                }
                                else if (tempArr[0] == "开机时间")
                                {
                                    sheet.Cells[rowNum, 3] = tempArr[3];
                                }
                            }
                        }
                        else //日期不一样
                        {
                            //计算运行时长

                            rowNum++;           //新起一行数据
                            sheet.Cells[rowNum, 1] = tempArr[1];  //日期
                            sheet.Cells[rowNum, 2] = tempArr[2];  //星期
                            if (tempArr[0] == "关机时间")
                            {
                                sheet.Cells[rowNum, 4] = tempArr[3];
                            }
                            else if (tempArr[0] == "开机时间")
                            {
                                sheet.Cells[rowNum, 3] = tempArr[3];
                            }
                        }
                    }
                    strLine = sr.ReadLine();
                }
                Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(sheet.Cells[1, 3], sheet.Cells[rowNum, 4]);
                range.NumberFormat = "hh:mm:ss";//设置时间格式hh:mm:ss  
                //RowHeight   "1:1"表示第一行, "1:2"表示,第一行和第二行 
                ((Excel.Range)sheet.Rows["1:20", System.Type.Missing]).RowHeight = 30;
                //ColumnWidth "A:B"表示第一列和第二列, "A:A"表示第一列
                ((Excel.Range)sheet.Columns["A:F", System.Type.Missing]).ColumnWidth = 15;
                //设置字体大小
                Excel.Range excelRange = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[rowNum, 5]);
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
