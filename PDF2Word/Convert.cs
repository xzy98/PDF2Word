using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire;
using Spire.Doc;
using Spire.Pdf;
using System.Windows.Forms;

namespace PDF2Word
{
    class Convert
    {
        public static bool PDF2Word(string src, string dec)
        {
            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(src);
            doc.SaveToFile(dec, Spire.Pdf.FileFormat.DOCX);
            //System.Diagnostics.Process.Start("图文版丽江旅游攻略大全.doc");
            MessageBox.Show("转换完成");
            return true;
        }
        public static bool Word2PDF(string src, string dec)
        {
            //PdfDocument doc = new PdfDocument();
            Document doc = new Document();
            doc.LoadFromFile(src);
            doc.SaveToFile(dec, Spire.Doc.FileFormat.PDF);
            //System.Diagnostics.Process.Start("图文版丽江旅游攻略大全.doc");
            MessageBox.Show("转换完成");
            return true;
        }

    }
}
