using System;
using System.IO;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Saving;

namespace Doc2Picture
{
    class DOC
    {
        public static void ToPDF()
        {
            Document document = new Document(@"D:\00.doc");
            document.Save(@"D:\00.pdf", SaveFormat.Pdf);
        }

        public static void ToJPG()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "文档文件|*.doc|文档文件|*.docx"
            };
            DialogResult dialogResult = openFileDialog.ShowDialog();
            if (dialogResult != DialogResult.OK) return;
            ///
            Document doc = new Document(openFileDialog.FileName);
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                PrettyFormat = true,
                UseHighQualityRendering = true,
                Resolution = 72
            };
            for (int i = 0; i < doc.PageCount; i++)
            {
                imageOptions.PageIndex = i;
                doc.Save(Path.Combine(Path.GetDirectoryName(openFileDialog.FileName), Path.GetFileNameWithoutExtension(openFileDialog.FileName) + "_" + i.ToString() + ".jpg"), imageOptions);
            }
        }

    }
}

