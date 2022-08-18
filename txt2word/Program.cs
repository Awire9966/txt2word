using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;
namespace txt2word
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Application App = new Application();
            Document Doc = App.Documents.Add();
            Doc.BuiltInDocumentProperties["Author"].Value = "txt2word by Awire9966 on github.";
            
            
            StringBuilder stringBuilder = new StringBuilder();
            using (FileStream fs = File.Open(args[0], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (BufferedStream bs = new BufferedStream(fs))
            using (StreamReader sr = new StreamReader(bs))
            {
                Doc.Content.Text = sr.ReadToEnd();
                Doc.SaveAs(args[1]);
            }
           
        }
    }
}
