using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace TxtToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            TxtToExcel("E:\\git\\01.开关机记录\\01.开关机数据\\开关机数据.log", "E:\\git\\TxtToExcel\\TxtToExcel\\bin\\Debug\\开关机数据.log");
            System.Threading.Thread.Sleep(8000);
            this.Close();
        }
        public static void TxtToExcel(string TxtFilePath,string OutputPath)
        {
            try
            {
                StreamReader sr = new StreamReader(TxtFilePath);
                string strLine = sr.ReadLine();
                sr.Close();
                //将指定内容写入指定文件
                StreamWriter sw = new StreamWriter(OutputPath);
                sw.WriteLine(strLine);
                sw.Close();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
}
