namespace Doc2Picture
{
    public static class CDR
    {
        /// <summary>
        /// CDR转PNG
        /// </summary>
        /// <param name="cdr_path">CDR位置</param>
        /// <param name="png_path">PNG位置</param>
        public static void ToPNG(string cdr_path, string png_path)
        {
            CorelDRAW.Application coreldraw = new CorelDRAW.Application();
            coreldraw.OpenDocument(cdr_path, 1);
            //System.Windows.Forms.MessageBox.Show(coreldraw.w
            coreldraw.ActiveDocument.ExportBitmap(
                png_path,
                VGCore.cdrFilter.cdrPNG,
                VGCore.cdrExportRange.cdrCurrentPage,
                VGCore.cdrImageType.cdrRGBColorImage,
                0, 0, 72, 72,
                VGCore.cdrAntiAliasingType.cdrNoAntiAliasing,
                false,
                true,
                true,
                false,
                VGCore.cdrCompressionType.cdrCompressionNone, null
                ).Finish();
            coreldraw.ActiveDocument.Close();
            coreldraw.Quit();
        }

        /// <summary>
        /// CDR转JPG
        /// </summary>
        /// <param name="cdr_path">CDR位置</param>
        /// <param name="jpg_path">JPG位置</param>
        public static void ToJPG(string cdr_path, string jpg_path)
        {
            CorelDRAW.Application coreldraw = new CorelDRAW.Application();
            coreldraw.OpenDocument(cdr_path, 1);
            //System.Windows.Forms.MessageBox.Show(coreldraw.StructFontProperties.ToString());
            coreldraw.ActiveDocument.ExportBitmap(
                jpg_path,
                VGCore.cdrFilter.cdrJPEG,
                VGCore.cdrExportRange.cdrCurrentPage,
                VGCore.cdrImageType.cdrRGBColorImage,
                0, 0, 72, 72,
                VGCore.cdrAntiAliasingType.cdrNoAntiAliasing,
                false,
                true,
                true,
                false,
                VGCore.cdrCompressionType.cdrCompressionNone, null
                ).Finish();
            coreldraw.ActiveDocument.Close();
            coreldraw.Quit();
        }

        /// <summary>
        /// CDR转PDF
        /// </summary>
        /// <param name="cdr_path">CDR位置</param>
        /// <param name="pdf_path">PDF位置</param>
        public static void ToPDF(string cdr_path, string pdf_path) {
            CorelDRAW.Application coreldraw = new CorelDRAW.Application();
            coreldraw.OpenDocument(cdr_path, 1);
            coreldraw.ActiveDocument.PublishToPDF(pdf_path);
            coreldraw.ActiveDocument.Close();
            coreldraw.Quit();
        }
    }
}
