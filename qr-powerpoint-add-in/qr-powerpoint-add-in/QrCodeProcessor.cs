using IronBarCode;

namespace qr_powerpoint_add_in
{
    public class QrCodeProcessor
    {
        public static string ConvertUrlToQrCode(string url)
        {
            return QRCodeWriter.CreateQrCode(url, 500, QRCodeWriter.QrErrorCorrectionLevel.Medium).ToDataUrl();
        }
    }
}
