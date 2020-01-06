
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RenameSheet
{
    public partial class Form1 : Form
    {
        FileInfo[] files;
        StringBuilder successStr=new StringBuilder();
        StringBuilder errorStr= new StringBuilder();
        public Form1()
        {
            InitializeComponent();
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            files=GetAllFiles("请选择文件夹");

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        public  FileInfo[] GetAllFiles(string prompt)
        {
            FileInfo[] result = null;
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            folderDialog.Description = prompt;
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = folderDialog.SelectedPath;
                DirectoryInfo theFolder = new DirectoryInfo(foldPath);
                result = theFolder.GetFiles();
            }
            textBox1.Text = folderDialog.SelectedPath + "\\";
            return result;
        }


        private void button2_Click(object sender, EventArgs e)
        {

            int i = 0;
            foreach (FileInfo item in files)
            {
                string name = item.Name;
                string fullName = item.FullName;

                if (fullName.IndexOf(".xlsx") > 0)
                {
                    //2007版本
                    try
                    {
                        Workbook workbook = new Workbook();
                        workbook.LoadFromFile(fullName);
                        Worksheet worksheet = workbook.Worksheets[0];
                        worksheet.Name = name.Replace(".xlsx", "");
                        workbook.Save();
                        Console.WriteLine(i++);
                        Console.WriteLine(name);
                        successStr.Append(name + "\n");
                    }
                    catch (Exception ex)
                    {
                        errorStr.Append(name + "\n");
                        Console.WriteLine(ex.Message);
                    }
                    richTextBox1.Text = successStr.ToString();
                    richTextBox2.Text = errorStr.ToString();
                }
                


            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
