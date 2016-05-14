using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointSlideIndicator_2010
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            this.Application.SlideShowNextClick += Application_SlideShowNextClick;
        }

        private void Application_SlideShowNextClick(PowerPoint.SlideShowWindow Wn, PowerPoint.Effect nEffect)
        {
            var slide = Wn.View.Slide;

            var slideMax = 0;
            foreach (PowerPoint.Slide s in Wn.Presentation.Slides)
                slideMax += s.SlideShowTransition.Hidden.Equals(Office.MsoTriState.msoFalse) ? 1 : 0;
            
            try
            {
                var shapes = Wn.View.Slide.Shapes.Range("indicator_obj");
                var indicator_width = shapes.Width;
                var slide_num_obj = Wn.View.Slide.Shapes.Range("スライド番号プレースホルダー 3");
                var left = slide_num_obj.Left;
                // all - title
                var delta = (left - indicator_width) / (slideMax - 2);
                // slide start at 2
                shapes.Left = (slide.SlideNumber - 2) * delta;
            }
            catch (Exception ex)
            {

            }

        }

        #endregion
    }
}
