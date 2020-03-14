using System;
using System.Windows.Forms;

namespace Doc2Picture
{
    public partial class HomeForm : Form
    {
        public HomeForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DOC.ToJPG();
        }
    }
}
