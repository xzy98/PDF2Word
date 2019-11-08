using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Spire.Doc;
using Spire.Pdf;
using System.IO;

namespace PDF2Word
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.ShowDialog();
            if (openFile.FileName == null)
            {
                MessageBox.Show("请选择文件");
                return;
            }
            string src = openFile.FileName;
            

            textBox1.Text = src;


        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
           
            if (path.SelectedPath == null)
            {
                MessageBox.Show("请选择目录");
                return;
            }
            string src = path.SelectedPath;

            textBox2.Text = src;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("未选择文件或目录");
                return;
            }
            string src = textBox1.Text;
            string fileName = @"\" + System.IO.Path.GetFileName(src).Split('.')[0];
            string path = textBox2.Text + fileName;
            

            
            if (System.IO.Path.GetExtension(src)==".pdf")
            {
                
                Convert.PDF2Word(src, path+".docx");
            }
            if (System.IO.Path.GetExtension(src) == ".doc"|| System.IO.Path.GetExtension(src) == ".docx")
            {
                
                Convert.Word2PDF(src, path+".pdf");
            }
        }
    }
}
