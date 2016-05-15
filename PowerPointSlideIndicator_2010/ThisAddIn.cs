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
            // ToDo: 変更できるようなインターフェース
            const string INDICATOR_OBJECT = "indicator_obj";
            const string SLIDE_NUMBER_OBJECT = "スライド番号プレースホルダー 3";


            var slide = Wn.View.Slide;

            var visible_slide = Wn.Presentation.Slides.Cast<Microsoft.Office.Interop.PowerPoint.Slide>().Count(
                s => s.SlideShowTransition.Hidden.Equals(Office.MsoTriState.msoFalse));
            
            try
            {
                var indicator_obj = Wn.View.Slide.Shapes.Range(INDICATOR_OBJECT);
                var indicator_width = indicator_obj.Width;
                var slide_num_obj = Wn.View.Slide.Shapes.Range(SLIDE_NUMBER_OBJECT);
                var left = slide_num_obj.Left;
                // all - title
                var delta = (left - indicator_width) / (visible_slide - 2);
                // slide start at 2
                indicator_obj.Left = (slide.SlideNumber - 2) * delta;
            }
            catch (Exception ex)
            {

            }

        }

        #endregion
    }
}
