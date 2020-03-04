using System;
using System.IO;
using System.Windows.Forms;
using Aspose.Pdf;
namespace DocumentsToPicture
{
  public static  class PDF
    {
        public static void Tojpg()
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "pdf文件|*.pdf";
            var dialogResult = dialog.ShowDialog();
            if (dialogResult != DialogResult.OK)return;
            
            //和选择的文件并列创建一个目录
            string filePath = dialog.FileName;
            string directoryPath = filePath + "目录";
            //aspose许可证
            Aspose.Pdf.License l = new Aspose.Pdf.License();
            string licenseName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Aspose.Total.Product.Family.lic");
            l.SetLicense(licenseName);
            //定义Jpeg转换设备
            Document pdf = new Document(filePath);
            var device = new Aspose.Pdf.Devices.JpegDevice();
            int quality = int.Parse(this.comboBox1.SelectedItem.ToString());
            directoryPath += quality;
            Directory.CreateDirectory(directoryPath);
            //默认质量为100，设置质量的好坏与处理速度不成正比，甚至是设置的质量越低反而花的时间越长，怀疑处理过程是先生成高质量的再压缩
            device = new Aspose.Pdf.Devices.JpegDevice(quality);
            //遍历每一页转为jpg
            for (var i = 1; i <= pdf.Pages.Count; i++)
            {
                string filePathOutPut = Path.Combine(directoryPath, string.Format("{0}.jpg", i));
                FileStream fs = new FileStream(filePathOutPut, FileMode.OpenOrCreate);
                try
                {
                    device.Process(pdf.Pages[i], fs);
                    fs.Close();
                }
                catch (Exception ex)
                {
                    fs.Close();
                    File.Delete(filePathOutPut);
                }
            }

        }
}
