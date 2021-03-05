using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;

namespace ProductivityTools.PowerPointLanguage
{
    public partial class ChangeLanguage
    {
        private void ChangeLanguage_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void slideButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.ViewType == PpViewType.ppViewNormal)
                Globals.ThisAddIn.Application.ActiveWindow.Panes[2].Activate();

            var currentSlideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex;
            var slides = Globals.ThisAddIn.Application.ActivePresentation.Slides[currentSlideIndex];
            ChangeLanguageForSlide(slides);
        }

        private void wholePresentation_Click(object sender, RibbonControlEventArgs e)
        {
            var slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            foreach (Slide s in slides)
            {
                ChangeLanguageForSlide(s);
            }
        }

        private void ChangeLanguageForSlide(Slide s)
        {
            foreach (Shape shape in s.Shapes)
            {
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    if (shape.TextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        shape.TextFrame2.TextRange.LanguageID = Microsoft.Office.Core.MsoLanguageID.msoLanguageIDEnglishUS;
                    }
                }
                if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                {

                    foreach (Row row in shape.Table.Rows)
                    {
                        foreach (Cell cells in row.Cells)
                        {
                            if (cells.Shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                if (cells.Shape.TextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                {
                                    cells.Shape.TextFrame2.TextRange.LanguageID = Microsoft.Office.Core.MsoLanguageID.msoLanguageIDEnglishUS;
                                }
                            }
                        }
                    }
                }

                if (shape.HasSmartArt == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    foreach (Microsoft.Office.Core.SmartArtNode x in shape.SmartArt.AllNodes)
                    {
                        if (x.TextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            x.TextFrame2.TextRange.LanguageID = Microsoft.Office.Core.MsoLanguageID.msoLanguageIDEnglishUS;
                        }
                    }

                    //foreach (Row row in shape.Table.Rows)
                    //{
                    //    foreach (Cell cells in row.Cells)
                    //    {
                    //        if (cells.Shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    //        {
                    //            if (cells.Shape.TextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    //            {
                    //                cells.Shape.TextFrame2.TextRange.LanguageID = Microsoft.Office.Core.MsoLanguageID.msoLanguageIDEnglishUS;
                    //            }
                    //        }
                    //    }
                    //}
                }
            }

            foreach (Shape notesPage in s.NotesPage.Shapes)
            {
                if (notesPage.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                {
                    if (notesPage.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
                    {
                        notesPage.TextFrame2.TextRange.LanguageID = Microsoft.Office.Core.MsoLanguageID.msoLanguageIDEnglishUS;
                    }
                }
            }
        }
    }
}
