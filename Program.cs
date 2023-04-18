using System.Text;
using MimeKit;
using MimeKit.Tnef;

namespace WinMailDatExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var filename = args[0];
            string extractPath = Path.Join(
                Path.GetDirectoryName(filename),
                @"extract"
            );

            if (!Directory.Exists(extractPath)) {
                Directory.CreateDirectory(extractPath);
            }
            
            Encoding shiftJIS = Encoding.GetEncoding("Shift_JIS");
            Console.OutputEncoding = shiftJIS;
            
            var attachment = new MimePart("winmail", "dat") {
                Content = new MimeContent(File.OpenRead(filename), ContentEncoding.Default),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = "winmail.dat"
            };

            var part = new TnefPart();
            part.Content = attachment.Content;
            foreach (var pattachment in part.ExtractAttachments())
            {
                if (!pattachment.IsAttachment) {
                    using (var stream = File.Create(Path.Join(extractPath, "content.rtf"))) {
                        ((MimePart) pattachment).Content.DecodeTo(stream);
                    }
                }
                if (pattachment.ContentDisposition?.FileName != null) {
                    using (var stream = File.Create (Path.Join(extractPath, pattachment.ContentDisposition?.FileName))) {
                        ((MimePart) pattachment).Content.DecodeTo(stream);
                    }
                }
                
            }
            
            Console.WriteLine("Done! {0}", extractPath);
        }
    }
}
