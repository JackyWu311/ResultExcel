using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ResultExcel
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        //private Point selectionStart;
        //private Point selectionEnd;
        //private Rectangle selection;
        //private bool mouseDown;

        //private void GetSelectedTextBoxes()
        //{
        //    List<TextBox> selected = new List<TextBox>();

        //    foreach (Control c in Controls)
        //    {
        //        if (c is TextBox)
        //        {
        //            if (selection.IntersectsWith(c.Bounds))
        //            {
        //                selected.Add((TextBox)c);
        //            }
        //        }
        //    }

        //    // Replace with your input box
        //    MessageBox.Show("You selected " + selected.Count + " textbox controls.");
        //}
        //private void Form2_MouseDown(object sender, MouseEventArgs e)
        //{
        //    Console.WriteLine("1231321");
        //    selectionStart = PointToClient(MousePosition);
        //    mouseDown = true;
        //}

        //private void Form2_MouseMove(object sender, MouseEventArgs e)
        //{
        //    if (!mouseDown)
        //    {
        //        return;
        //    }

        //    selectionEnd = PointToClient(MousePosition);
        //    SetSelectionRect();

        //    Invalidate();
        //}

        //private void Form2_MouseUp(object sender, MouseEventArgs e)
        //{
        //    mouseDown = false;

        //    SetSelectionRect();
        //    Invalidate();

        //    GetSelectedTextBoxes();
        //}

        //private void Form2_Paint(object sender, PaintEventArgs e)
        //{
        //    //base.OnPaint(e);

        //    if (mouseDown)
        //    {
        //        using (Pen pen = new Pen(Color.Black, 1F))
        //        {
        //            pen.DashStyle = DashStyle.Dash;
        //            e.Graphics.DrawRectangle(pen, selection);
        //        }
        //    }
        //}
        //private void SetSelectionRect()
        //{
        //    int x, y;
        //    int width, height;

        //    x = selectionStart.X > selectionEnd.X ? selectionEnd.X : selectionStart.X;
        //    y = selectionStart.Y > selectionEnd.Y ? selectionEnd.Y : selectionStart.Y;

        //    width = selectionStart.X > selectionEnd.X ? selectionStart.X - selectionEnd.X : selectionEnd.X - selectionStart.X;
        //    height = selectionStart.Y > selectionEnd.Y ? selectionStart.Y - selectionEnd.Y : selectionEnd.Y - selectionStart.Y;

        //    selection = new Rectangle(x, y, width, height);
        //}

        private Point selectionStartF;
        private Point selectionEndF;
        private Rectangle selectionF;
        private bool mouseDownF;
        private void GetSelectedTextBoxes2()
        {
            List<TextBox> selected = new List<TextBox>();

            foreach (Control c in flowLayoutPanel1.Controls)
            {
                if (c is TextBox)
                {
                    if (selectionF.IntersectsWith(c.Bounds))
                    {
                        selected.Add((TextBox)c);
                    }
                }
            }
        }
            private void flowLayoutPanel1_MouseDown(object sender, MouseEventArgs e)
        {
            selectionStartF = e.Location;
            mouseDownF = true;
        }

        private void flowLayoutPanel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (!mouseDownF)
            {
                return;
            }

            selectionEndF = e.Location;
            SetSelectionRectF();

            Invalidate();
        }

        private void flowLayoutPanel1_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDownF = false;

            SetSelectionRectF();
            Invalidate();

            GetSelectedTextBoxes2();
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
            //base.OnPaint(e);

            if (mouseDownF)
            {
                Console.WriteLine("AAAAA");
                using (Pen pen = new Pen(Color.Black, 1F))
                {
                    pen.DashStyle = DashStyle.Dash;
                    e.Graphics.DrawRectangle(pen, selectionF);
                }
            }
        }
        private void SetSelectionRectF()
        {
            int x, y;
            int width, height;

            x = selectionStartF.X > selectionEndF.X ? selectionEndF.X : selectionStartF.X;
            y = selectionStartF.Y > selectionEndF.Y ? selectionEndF.Y : selectionStartF.Y;

            width = selectionStartF.X > selectionEndF.X ? selectionStartF.X - selectionEndF.X : selectionEndF.X - selectionStartF.X;
            height = selectionStartF.Y > selectionEndF.Y ? selectionStartF.Y - selectionEndF.Y : selectionEndF.Y - selectionStartF.Y;

            selectionF = new Rectangle(x, y, width, height);
        }
    }
}
