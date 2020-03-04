using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;

namespace DocumentsToPicture
{
    class DOC
    {
        public static void ToJpg()
        {
            Document doc = new Document(@"C:\Users\Administrator\Desktop\123\123.doc");
            doc.Save(@"C:\Users\Administrator\Desktop\temp.pdf", SaveFormat.Pdf);
        }
    }
}

