using SwDocumentMgr;
using System;
using System.Drawing;
using System.IO;

namespace CodeStackExample
{
    class Program
    {
        const string LICENSE_KEY = "[Document Manager License Key]";

        static void Main(string[] args)
        {
            var filePath = args[0];
            var outImgFilePath = args[1];

            var classFact = new SwDMClassFactory();

            var app = classFact.GetApplication(LICENSE_KEY);

            var docType = SwDmDocumentType.swDmDocumentUnknown;

            switch (Path.GetExtension(filePath).ToLower())
            {
                case ".sldprt":
                    docType = SwDmDocumentType.swDmDocumentPart;
                    break;

                case ".sldasm":
                    docType = SwDmDocumentType.swDmDocumentAssembly;
                    break;

                case ".slddrw":
                    docType = SwDmDocumentType.swDmDocumentDrawing;
                    break;
            }

            SwDmDocumentOpenError err;
            var doc = app.GetDocument(filePath,
                SwDmDocumentType.swDmDocumentPart, true, out err);

            if (doc != null)
            {
                var activeConfName = doc.ConfigurationManager.GetActiveConfigurationName();

                var conf = doc.ConfigurationManager.GetConfigurationByName(activeConfName) as ISwDMConfiguration14;

                SwDmPreviewError previewErr;
                var imgBytes = conf.GetPreviewPNGBitmapBytes(out previewErr) as byte[];

                if (previewErr == SwDmPreviewError.swDmPreviewErrorNone)
                {
                    using (var memStr = new MemoryStream(imgBytes))
                    {
                        memStr.Seek(0, SeekOrigin.Begin);
                        var img = Image.FromStream(memStr);
                        img.Save(outImgFilePath);
                    }
                }
                else
                {
                    Console.WriteLine($"Failed to extract preview from the document: {previewErr}");
                }
            }
            else
            {
                Console.WriteLine($"Failed to open the document: {err}");
            }
        }
    }
}