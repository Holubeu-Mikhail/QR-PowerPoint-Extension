using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace qr_powerpoint_add_in
{
    public partial class AddInProgram
    {
        private string _encodedUrl = null;
        private string _filename = "qr.png";
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += CreateImageClipboard;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void CreateImageClipboard(PowerPoint.Selection selection)
        {
            if (Clipboard.ContainsText())
            {
                string text = Clipboard.GetText();
                if(_encodedUrl != text && Regex.IsMatch(Clipboard.GetText(), @".*docs\.google\.com\/forms\/.*"))
                {
                    QrCodeProcessor.ConvertUrlToQrCode(text, _filename);
                    AddQrImageOnSlide(text);
                    File.Delete(_filename);
                }
            }
        }

        private void AddQrImageOnSlide(string url)
        {
            PowerPoint.Slide activeSlide = Globals.AddInProgram.Application.ActiveWindow.View.Slide;
            PowerPoint.Shape ppPicture = activeSlide.Shapes.AddPicture(_filename, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Application.ActivePresentation.PageSetup.SlideWidth - 300, Application.ActivePresentation.PageSetup.SlideHeight - 300);
            ppPicture.LinkFormat.SourceFullName = _filename;
            _encodedUrl = url;
        }

        void CreateImageOnShapeChanged(PowerPoint.Selection selection)
        {
            var range = selection.ShapeRange;
            foreach (var shape in range)
            {
                var sh = (PowerPoint.Shape)shape;
                if (sh.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    string text = sh.TextFrame.TextRange.Text;
                    if (Regex.IsMatch(text, @".*docs\.google\.com\/forms\/.*") && _encodedUrl != text)
                    {
                        QrCodeProcessor.ConvertUrlToQrCode(text, _filename);

                        PowerPoint.Slide activeSlide = Globals.AddInProgram.Application.ActiveWindow.View.Slide;
                        PowerPoint.Shape ppPicture = activeSlide.Shapes.AddPicture(_filename, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, 0, 0);
                        ppPicture.LinkFormat.SourceFullName = _filename;
                        _encodedUrl = text;

                        File.Delete(_filename);
                    }
                }
            }
        }

        private void CreateImageOnResize(PowerPoint.Shape shape)
        {
            var text = shape.TextFrame.TextRange.Text;
            if (shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                if (Regex.IsMatch(text, @".*docs\.google\.com\/forms\/.*"))
                {
                    QrCodeProcessor.ConvertUrlToQrCode(text, _filename);

                    PowerPoint.Slide activeSlide = Globals.AddInProgram.Application.ActiveWindow.View.Slide;
                    PowerPoint.Shape ppPicture = activeSlide.Shapes.AddPicture(_filename, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, 0, 0);
                    ppPicture.LinkFormat.SourceFullName = _filename;

                    File.Delete(_filename);
                }
            }
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