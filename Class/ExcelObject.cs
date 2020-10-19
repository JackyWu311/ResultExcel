using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ResultExcel.Class
{
    class ExcelObject
    {
        public string HtmlColumn { get; set; }
        public string ExcelFile;
        public string HtmlFile;
        public List<Block> blocks;

        public ExcelObject(string htmlcolumn,List<Block> blocks, string ExcelFile, string HtmlFile)
        {
            HtmlColumn = htmlcolumn;
            this.blocks = blocks;
            this.ExcelFile = ExcelFile;
            this.HtmlFile = HtmlFile;
        }
    }
}
