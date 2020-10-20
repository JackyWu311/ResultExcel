using System.Collections.Generic;

namespace ResultExcel.Class
{
    internal class ExcelObject
    {
        public string HtmlColumn { get; set; }
        public string ExcelFile;
        public string HtmlFile;
        public List<Block> blocks;

        public ExcelObject(string htmlcolumn, List<Block> blocks, string ExcelFile, string HtmlFile)
        {
            HtmlColumn = htmlcolumn;
            this.blocks = blocks;
            this.ExcelFile = ExcelFile;
            this.HtmlFile = HtmlFile;
        }
    }
}