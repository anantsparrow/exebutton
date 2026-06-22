using System;
using System.Drawing;
using System.Windows.Forms;

namespace PhishingReporterAddIn
{
    public partial class NotificationForm : Form
    {
        private Color _accentColor;
        private Color _buttonNormalColor = Color.FromArgb(46, 46, 62);
        private Color _buttonHoverColor = Color.FromArgb(62, 62, 78);

        public NotificationForm(bool isSuccess, string title, string message)
        {
            InitializeComponent();
            SetupForm(isSuccess, title, message);
        }

        private void SetupForm(bool isSuccess, string title, string message)
        {
            // Emerald Green for success, Flat Crimson Red for error
            _accentColor = isSuccess ? Color.FromArgb(46, 204, 113) : Color.FromArgb(231, 76, 60);

            this.pnlAccent.BackColor = _accentColor;
            this.lblTitle.Text = title;
            this.lblMessage.Text = message;

            // Setup unicode icons
            this.lblIcon.Text = isSuccess ? "✔" : "✘";
            this.lblIcon.ForeColor = _accentColor;

            // Configure button styling and hover effects
            this.btnOk.BackColor = _buttonNormalColor;
            this.btnOk.ForeColor = Color.White;
            this.btnOk.FlatStyle = FlatStyle.Flat;
            this.btnOk.FlatAppearance.BorderSize = 0;
            this.btnOk.Cursor = Cursors.Hand;

            this.btnOk.MouseEnter += BtnOk_MouseEnter;
            this.btnOk.MouseLeave += BtnOk_MouseLeave;

            // Register Paint event to draw a thin high-quality border around the form
            this.Paint += NotificationForm_Paint;
        }

        private void BtnOk_MouseEnter(object sender, EventArgs e)
        {
            this.btnOk.BackColor = _buttonHoverColor;
        }

        private void BtnOk_MouseLeave(object sender, EventArgs e)
        {
            this.btnOk.BackColor = _buttonNormalColor;
        }

        private void NotificationForm_Paint(object sender, PaintEventArgs e)
        {
            // Draw a subtle border inside the form boundaries
            using (Pen borderPen = new Pen(Color.FromArgb(45, 45, 58), 1))
            {
                e.Graphics.DrawRectangle(borderPen, 0, 0, this.Width - 1, this.Height - 1);
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
