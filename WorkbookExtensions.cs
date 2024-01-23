using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

namespace AsposeCells.DigitalSigning.Test
{
    public class WorkbookExtensions
    {
        public static Workbook CreateWorkbook()
        {
            var loadOptions = new LoadOptions(LoadFormat.Auto)
            {
                MemorySetting = MemorySetting.MemoryPreference
            };

            const string path = @"Contents\Book1.xlsm";

            var readText = File.ReadAllBytes(path);
            var stream = new MemoryStream(readText);

            var workbook = new Workbook(stream, loadOptions);
            return workbook;
        }

        public static DigitalSignature GetDigitalSignature()
        {
            const string comment = "Signing Digital Signature using Aspose.Cells";

            const string pfxPath = @"Contents\friendly-cert2.pfx";

            var cert = new X509Certificate2(pfxPath, "mypassword", X509KeyStorageFlags.MachineKeySet);

            var digitalSignature =
                new DigitalSignature
                    (cert, comment, DateTime.UtcNow);
            return digitalSignature;
        }
    }

    public class MyStream
    {
        private const string FileName = "Test.xlsm";

        public static void Main(byte[] content)
        {
            if (File.Exists(FileName))
            {
                Console.WriteLine($"{FileName} already exists!");
                return;
            }

            using (var fs = new FileStream(FileName, FileMode.CreateNew))
            {
                fs.Write(content, 0, content.Length);
            }

            using (var fs = new FileStream(FileName, FileMode.Open, FileAccess.Read))
            {
                using var r = new BinaryReader(fs);
                for (var i = 0; i < 11; i++)
                {
                    Console.WriteLine(r.ReadInt32());
                }
            }
        }
    }
}