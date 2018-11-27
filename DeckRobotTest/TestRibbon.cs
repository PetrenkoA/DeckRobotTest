using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System.Drawing;

namespace DeckRobotTest
{
    public partial class TestRibbon
    {
        private void NewTestRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        static float LeftShift = 10;

        private void btn_test_Click(object sender, RibbonControlEventArgs e)
        {
            Application thisApp = Globals.ThisAddIn.Application;
            foreach (Slide slide in thisApp.ActiveWindow.Selection.SlideRange)
            {
                foreach(Shape elem in slide.Shapes)
                {
                    if (elem.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && 
                        elem.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue &&
                        elem.Left + elem.Width + LeftShift <= thisApp.ActiveWindow.Presentation.PageSetup.SlideWidth)
                    {
                        elem.Left += LeftShift;
                        elem.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                        elem.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Red);
                    }
                }
            }
        }

    }
}
