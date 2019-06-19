using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace Blackboard_Quiz_Form
{
    public partial class ThisDocument
    {
        private List<object> Questions = new List<object>();
        public Word.ContentControl repeatingDistractor;
        public Word.ContentControl repeatingQuestion;
        public int i = 1;

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            Word.Range distractorRange = richTextContentControl1.Range;
            distractorRange.Select();
            Word.Range w = this.Application.Selection.Range;
            repeatingDistractor = this.ContentControls.Add(Word.WdContentControlType.wdContentControlRepeatingSection, w);
            repeatingDistractor.RepeatingSectionItems[1].InsertItemAfter();
            Word.Range questionRange = buildingBlockGalleryContentControl1.Range;
            questionRange.Select();
            Word.Range r = this.Application.Selection.Range;
            repeatingQuestion = this.ContentControls.Add(Word.WdContentControlType.wdContentControlRepeatingSection, r);
            //repeatingQuestion.RepeatingSectionItems[1].InsertItemAfter();
            repeatingQuestion.RepeatingSectionItemTitle = "Question " + i;
            i++;
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
            Console.Write(Questions);
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.buildingBlockGalleryContentControl1.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.buildingBlockGalleryContentControl1_Entering);
            this.buildingBlockGalleryContentControl1.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler(this.buildingBlockGalleryContentControl1_Exiting);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion

        private void buildingBlockGalleryContentControl1_Entering(object sender, ContentControlEnteringEventArgs e)
        {

        }

        private void buildingBlockGalleryContentControl1_Exiting(object sender, ContentControlExitingEventArgs e)
        {
            Word.ContentControl FirstQuestion = SelectContentControlsByTag("question")[1];
            i = 1;
            foreach (Word.ContentControl question in FirstQuestion.Range.ContentControls)
            {
                if (question.Tag == "question")
                {
                    question.Title = "Question " + i;
                    i++;
                }
            }
        }
    }
}
