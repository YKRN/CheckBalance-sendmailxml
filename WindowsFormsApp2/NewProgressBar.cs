using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
   
        public class NewProgressBar : ProgressBar
        {
            public NewProgressBar()
            {
                this.SetStyle(ControlStyles.UserPaint, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
            LinearGradientBrush brush = null;
            Rectangle rec = e.ClipRectangle;

                rec.Width = (int)(rec.Width * ((double)Value / Maximum)) - 4;
                if (ProgressBarRenderer.IsSupported)
                    ProgressBarRenderer.DrawHorizontalBar(e.Graphics, e.ClipRectangle);
                rec.Height = rec.Height - 4;
                e.Graphics.FillRectangle(Brushes.Red, 2, 2, rec.Width, rec.Height);
            brush = new LinearGradientBrush(rec, this.ForeColor, this.BackColor, LinearGradientMode.Vertical);
        }
        }
    
}
