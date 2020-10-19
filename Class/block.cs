using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ResultExcel.Class
{
    public class Block
    {
        public string HtmlNo { get; set; }
        public string Success { get; set; }
        public string SuccessContent { get; set; }
        public string Fail { get; set; }
        public string FailContent { get; set; }
        public string Sheet { get; set; }
        public string Cell { get; set; }
        public string Comment { get; set; }
        public string Note { get; set; }
        public string NoteContent { get; set; }
        public string toString()
        {
            return HtmlNo + '\n' + Success + '\n' + SuccessContent + '\n' + Fail + '\n' + FailContent + '\n' + Sheet + '\n' + Cell + '\n' + Comment + '\n' + Note + '\n' + NoteContent;
        }
    }
}
