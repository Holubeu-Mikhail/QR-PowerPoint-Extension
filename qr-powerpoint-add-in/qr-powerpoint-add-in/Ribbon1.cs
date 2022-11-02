using Microsoft.Office.Tools.Ribbon;
using System.Drawing;

namespace qr_powerpoint_add_in
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string text = editBox1.Text;
            QrCodeProcessor.ConvertUrlToQrCode(text);
            label1.Label = text;
            //QrCodeProcessor.ConvertUrlToQrCode(text);
        }
    }
}
