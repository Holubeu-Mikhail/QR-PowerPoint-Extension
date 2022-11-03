using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;
using IronBarCode;

namespace qr_powerpoint_add_in
{
    public partial class AddInProgram
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.AfterShapeSizeChange +=
                new PowerPoint.EApplication_AfterShapeSizeChangeEventHandler(CreateImageFromLink);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void CreateImageFromLink(PowerPoint.Shape Sld)
        {
            var text = Sld.TextFrame.TextRange.Text;
            if (Sld.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                if (Regex.IsMatch(text, @".*docs\.google\.com\/forms\/.*"))
                {
                    string imageUrl = QrCodeProcessor.ConvertUrlToQrCode(text);

                    PowerPoint.Slide activeSlide = Globals.AddInProgram.Application.ActiveWindow.View.Slide;
                    PowerPoint.Shape ppPicture = activeSlide.Shapes.AddPicture(imageUrl, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, 0, 0);
                    ppPicture.LinkFormat.SourceFullName = imageUrl;
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
