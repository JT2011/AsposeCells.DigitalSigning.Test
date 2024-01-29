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

            const string path = @"Contents\workbooks\test-macros-file.xlsm";

            var readText = File.ReadAllBytes(path);
            var stream = new MemoryStream(readText);

            var workbook = new Workbook(stream, loadOptions);
            return workbook;
        }

        public static DigitalSignature GetDigitalSignature()
        {
            const string comment = "Signing Digital Signature using Aspose.Cells";

            const string devWorkbookHcrcentralComPfxPath = @"Contents\certs\dev-workbook.hcrcentral.com.pfx";
            const string friendlyCert2PfxPath = @"Contents\certs\friendly-cert2.pfx";

            var devX509Certificate2 = new X509Certificate2(devWorkbookHcrcentralComPfxPath, string.Empty, X509KeyStorageFlags.MachineKeySet);
            var devX509Certificatev3 = new X509Certificate(devWorkbookHcrcentralComPfxPath, string.Empty, X509KeyStorageFlags.MachineKeySet);
            var friendlyX509Certificate2 = new X509Certificate2(friendlyCert2PfxPath, "mypassword", X509KeyStorageFlags.MachineKeySet);

            var vertBytes = devX509Certificatev3.GetRawCertData();

            var digitalSignature =
                new DigitalSignature
                    (vertBytes, string.Empty, comment, DateTime.UtcNow);
            //var digitalSignature =
            //    new DigitalSignature
            //        (friendlyX509Certificate2, comment, DateTime.UtcNow);
            return digitalSignature;
        }
    }

    public class MyStream
    {
        private const string FileName = @"Contents\Digitally-Signed-Workbook.xlsm";

        public static void Main(byte[] content)
        {
            if (File.Exists(FileName))
            {
                Console.WriteLine($"{FileName} already exists!");
                return;
            }

            //writes to <root>\AsposeCells-DigitalSigning\AsposeCells.DigitalSigning.Test\bin\Debug\net8.0\Contents
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