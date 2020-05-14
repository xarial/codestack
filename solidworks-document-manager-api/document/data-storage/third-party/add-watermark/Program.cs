using SolidWorks.Interop.swdocumentmgr;
using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using System.Xml.Serialization;
using ThirdPartyStorage;

namespace CodeStack
{
    public class Watermark
    {
        public string CompanyName { get; set; }
        public string SignedBy { get; set; }
        public DateTime SignedOn { get; set; }
    }

    class Program
    {
        private const string DM_LIC_KEY = "<Your DM License Key>";
        private const string STORAGE_NAME = "_CodeStackWatermark";

        static void Main(string[] args)
        {
            var filePath = args[0];

            var isWriting = args[1] == "-w";

            var docType = GetDocumentType(filePath);
            
            var dmApp = ConnectoToDm(DM_LIC_KEY);

            SwDmDocumentOpenError err;
            var doc = dmApp.GetDocument(filePath, docType, !isWriting, out err) as SwDMDocument19;

            if (doc != null)
            {
                var stream = doc.Get3rdPartyStorage(STORAGE_NAME, isWriting) as IStream;

                try
                {
                    if (isWriting)
                    {
                        AddWatermark(args[2], stream);
                    }
                    else
                    {
                        ReadWatermark(stream);
                    }
                }
                catch
                {
                    throw;
                }
                finally
                {
                    doc.Release3rdPartyStorage(STORAGE_NAME);

                    if (isWriting)
                    {
                        doc.Save();
                    }

                    doc.CloseDoc();
                }
            }
            else
            {
                throw new NullReferenceException($"Failed to open the document: {err}");
            }
        }

        private static void ReadWatermark(IStream stream)
        {
            if (stream != null)
            {
                using (var comStream = new ComStream(stream, false, false))
                {
                    var ser = new XmlSerializer(typeof(Watermark));
                    var wm = ser.Deserialize(comStream) as Watermark;

                    Console.WriteLine($"Company Name: {wm.CompanyName}");
                    Console.WriteLine($"Signed By: {wm.SignedBy}");
                    Console.WriteLine($"Signed On: {wm.SignedOn}");
                }
            }
            else
            {
                Console.WriteLine("No watermark");
            }
        }

        private static void AddWatermark(string companyName, IStream stream)
        {
            var wm = new Watermark()
            {
                CompanyName = companyName,
                SignedBy = Environment.UserName,
                SignedOn = DateTime.Now
            };

            using (var comStream = new ComStream(stream, true, false))
            {
                var ser = new XmlSerializer(wm.GetType());
                ser.Serialize(comStream, wm);
            }

            Console.WriteLine("Watermark is added");
        }

        private static SwDMApplication ConnectoToDm(string licKey)
        {
            var classFact = new SwDMClassFactory();
            var docMgr = classFact.GetApplication(licKey) as SwDMApplication;

            return docMgr;
        }

        private static SwDmDocumentType GetDocumentType(string filePath)
        {
            var docType = SwDmDocumentType.swDmDocumentUnknown;

            switch (Path.GetExtension(filePath).ToUpper())
            {
                case ".SLDPRT":
                    docType = SwDmDocumentType.swDmDocumentPart;
                    break;
                case ".SLDASM":
                    docType = SwDmDocumentType.swDmDocumentAssembly;
                    break;
                case ".SLDDRW":
                    docType = SwDmDocumentType.swDmDocumentDrawing;
                    break;
                default:
                    throw new NotSupportedException("File type not supported");

            }

            return docType;
        }
    }
}
