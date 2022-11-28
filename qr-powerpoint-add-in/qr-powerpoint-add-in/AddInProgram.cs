using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MouseKeyboardActivityMonitor;
using MouseKeyboardActivityMonitor.WinApi;

namespace qr_powerpoint_add_in
{
    public partial class AddInProgram
    {
        private const string FILENAME = "qr.png";
        private const string FORMS_REGEX = @".*docs\.google\.com\/forms\/.*";

        private KeyboardHookListener _keyboardListener;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _keyboardListener = new KeyboardHookListener(new AppHooker());
            _keyboardListener.Enabled = true;
            _keyboardListener.KeyDown += CreateImageClipboard;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _keyboardListener.KeyDown -= CreateImageClipboard;
            _keyboardListener.Enabled = false;
            _keyboardListener = null;
        }

        private void CreateImageClipboard(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                if (Clipboard.ContainsText())
                {
                    var text = Clipboard.GetText();
                    if (Regex.IsMatch(text, FORMS_REGEX))
                    {
                        QrCodeProcessor.ConvertUrlToQrCode(text, FILENAME);
                        AddQrImageOnSlide();
                        if (!File.Exists(FILENAME)) return;
                        File.Delete(FILENAME);
                    }
                }
            }
        }

        private void AddQrImageOnSlide()
        {
            PowerPoint.Slide activeSlide = Globals.AddInProgram.Application.ActiveWindow.View.Slide;
            PowerPoint.Shape ppPicture = activeSlide.Shapes.AddPicture
                (FILENAME, Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                                Application.ActivePresentation.PageSetup.SlideWidth - 300, Application.ActivePresentation.PageSetup.SlideHeight - 300);

            activeSlide.Shapes.PasteSpecial();


            PowerPoint.Shape textbox = activeSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 100, 100);


            ppPicture.LinkFormat.SourceFullName = FILENAME;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}