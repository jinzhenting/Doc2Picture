using System;
using System.IO;
using System.Windows.Forms;
namespace Doc2Picture
{
    public static class PDF
    {
        public static void ToJPG()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "PDF|*.pdf"
            };
            DialogResult dialogResult = openFileDialog.ShowDialog();
            if (dialogResult != DialogResult.OK) return;
        }
    }
}
