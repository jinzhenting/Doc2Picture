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
            CDR.Topng(@"\\nrL\Vector\Today\1.cdr", @"\\nrL\Vector\Today\1.png");
            CDR.Tojpg(@"\\nrL\Vector\Today\1.cdr", @"\\nrL\Vector\Today\1.jpg");
            CDR.Topdf(@"\\nrL\Vector\Today\1.cdr", @"\\nrL\Vector\Today\1.pdf");
        }
    }
}
