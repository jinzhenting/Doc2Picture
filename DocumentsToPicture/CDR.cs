namespace DocumentsToPicture
{
    public static class CDR
    {
        /// <summary>
        /// CDR转PNG
        /// </summary>
        /// <param name="cdr_path">CDR位置</param>
        /// <param name="png_path">PNG位置</param>
        public static void Topng(string cdr_path, string png_path)
        {
            Corel.Interop.CorelDRAW.Application coreldraw = new Corel.Interop.CorelDRAW.Application();
            coreldraw.OpenDocument(cdr_path, 1);
            //System.Windows.Forms.MessageBox.Show(coreldraw.w
            coreldraw.ActiveDocument.ExportBitmap(
                png_path,
                Corel.Interop.VGCore.cdrFilter.cdrPNG,
                Corel.Interop.VGCore.cdrExportRange.cdrCurrentPage,
                Corel.Interop.VGCore.cdrImageType.cdrRGBColorImage,
                0, 0, 72, 72,
                Corel.Interop.VGCore.cdrAntiAliasingType.cdrNoAntiAliasing,
                false,
                true,
                true,
                false,
                Corel.Interop.VGCore.cdrCompressionType.cdrCompressionNone, null
                ).Finish();
            coreldraw.ActiveDocument.Close();
            coreldraw.Quit();
        }

        /// <summary>
        /// CDR转JPG
        /// </summary>
        /// <param name="cdr_path">CDR位置</param>
        /// <param name="jpg_path">JPG位置</param>
        public static void Tojpg(string cdr_path, string jpg_path)
        {
            Corel.Interop.CorelDRAW.Application coreldraw = new Corel.Interop.CorelDRAW.Application();
            coreldraw.OpenDocument(cdr_path, 1);
            //System.Windows.Forms.MessageBox.Show(coreldraw.StructFontProperties.ToString());
            coreldraw.ActiveDocument.ExportBitmap(
                jpg_path,
                Corel.Interop.VGCore.cdrFilter.cdrJPEG,
                Corel.Interop.VGCore.cdrExportRange.cdrCurrentPage,
                Corel.Interop.VGCore.cdrImageType.cdrRGBColorImage,
                0, 0, 72, 72,
                Corel.Interop.VGCore.cdrAntiAliasingType.cdrNoAntiAliasing,
                false,
                true,
                true,
                false,
                Corel.Interop.VGCore.cdrCompressionType.cdrCompressionNone, null
                ).Finish();
            coreldraw.ActiveDocument.Close();
            coreldraw.Quit();
        }

        /// <summary>
        /// CDR转PDF
        /// </summary>
        /// <param name="cdr_path">CDR位置</param>
        /// <param name="pdf_path">PDF位置</param>
        public static void Topdf(string cdr_path, string pdf_path) {
            Corel.Interop.CorelDRAW.Application coreldraw = new Corel.Interop.CorelDRAW.Application();
            coreldraw.OpenDocument(cdr_path, 1);
            coreldraw.ActiveDocument.PublishToPDF(pdf_path);
            coreldraw.ActiveDocument.Close();
            coreldraw.Quit();
        }
    }
}
