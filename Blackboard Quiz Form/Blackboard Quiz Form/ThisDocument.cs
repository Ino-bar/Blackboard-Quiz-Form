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
        
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            int i = 1;
            int questions = SelectContentControlsByTag("question").Count;
            foreach(object question in SelectContentControlsByTag("question"))
            {
                SelectContentControlsByTag("question")[i].Title = "Question " + i;
                i++;
            }
            Console.Write(questions);
            Console.Write(Questions);
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
            this.plainTextContentControl1.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.plainTextContentControl1_Entering);
            this.plainTextContentControl1.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler(this.plainTextContentControl1_Exiting);
            this.richTextContentControl1.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.richTextContentControl1_Entering);
            this.richTextContentControl2.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.richTextContentControl2_Entering);
            this.richTextContentControl3.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler(this.richTextContentControl3_Exiting);
            this.plainTextContentControl2.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.plainTextContentControl2_Entering);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion

        private void richTextContentControl2_Entering(object sender, ContentControlEnteringEventArgs e)
        {

        }

        private void plainTextContentControl2_Entering(object sender, ContentControlEnteringEventArgs e)
        {

        }

        private void plainTextContentControl1_Entering(object sender, ContentControlEnteringEventArgs e)
        {

        }

        private void richTextContentControl1_Entering(object sender, ContentControlEnteringEventArgs e)
        {

        }

        private void richTextContentControl3_Exiting(object sender, ContentControlExitingEventArgs e)
        {

        }

        private void plainTextContentControl1_Exiting(object sender, ContentControlExitingEventArgs e)
        {
            int i = 1;
            foreach (object question in SelectContentControlsByTag("question"))
            {
                SelectContentControlsByTag("question")[i].Title = "Question " + i;
                i++;
                //SelectContentControlsByTag("question").C
            }           
        }
        private void questionAdded(object sender, ContentControlAddedEventArgs e)
        {
        }
    }
}
