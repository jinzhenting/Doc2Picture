using System;
using System.Windows.Forms;

namespace DocumentsToPicture
{
    public partial class Index : Form
    {
        public Index()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DOC.ToJPG();
        }
    }
}
