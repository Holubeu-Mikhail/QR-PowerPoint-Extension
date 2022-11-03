using IronBarCode;

namespace qr_powerpoint_add_in
{
    public static class QrCodeProcessor
    {
        public static void ConvertUrlToQrCode(string url, string filename)
        {
            string logoFilename = "logo.png";
            QRCodeLogo logo = new QRCodeLogo(logoFilename, 75);
            QRCodeWriter.CreateQrCodeWithLogo(url, logo, 250).SaveAsPng(filename);
        }
    }
}
