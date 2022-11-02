using IronBarCode;

namespace qr_powerpoint_add_in
{
    public class QrCodeProcessor
    {
        public static void ConvertUrlToQrCode(string url)
        {
            GeneratedBarcode qrCode = QRCodeWriter.CreateQrCode(url);

            qrCode.SaveAsHtmlFile("qrCode.html");

            System.Diagnostics.Process.Start("qrCode.html");
        }
    }
}
