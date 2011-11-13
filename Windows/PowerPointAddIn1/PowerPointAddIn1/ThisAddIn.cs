using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;


namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Application oPPT;
            Microsoft.Office.Interop.PowerPoint.Presentations objPresSet;
            Microsoft.Office.Interop.PowerPoint.Presentation objPres;
            Microsoft.Office.Interop.PowerPoint.Slide slide;

            // the location of your powerpoint presentation
            string strPres = @"/Users/Matt/Desktop/sample.pptx";

            // create an instance of the PowerPoint
            oPPT = new Microsoft.Office.Interop.PowerPoint.Application();

            // show PowerPoint to the user
            oPPT.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

            objPresSet = oPPT.Presentations;

            // open the presentation
            objPres = objPresSet.Open(strPres, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);

            objPres.SlideShowSettings.Run();

            // advance the slideshow three slides
            objPres.SlideShowWindow.View.Next();
            objPres.SlideShowWindow.View.Next();
            objPres.SlideShowWindow.View.Next();

            // initialize slide object for first slide
            slide = objPres.Slides._Index(1);

            // export image of that slide
            slide.Export(@"\Users\Matt\Desktop\Slides\slide.png", "png", 320, 240);

            // extract notes from slide and save to text file
            string s = slide.NotesPage.Shapes[2].TextFrame.TextRange.Text;

            System.IO.File.WriteAllText(@"\Users\Matt\Desktop\Slides\note.txt", s);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
