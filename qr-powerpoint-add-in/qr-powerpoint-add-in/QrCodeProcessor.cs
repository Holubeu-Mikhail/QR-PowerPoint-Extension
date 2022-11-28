using IronBarCode;

namespace qr_powerpoint_add_in
{
    public static class QrCodeProcessor
    {
        public static void ConvertUrlToQrCode(string url, string filename)
        {
            QRCodeWriter.CreateQrCode(url, 250).SaveAsPng(filename);
        }
    }
}
