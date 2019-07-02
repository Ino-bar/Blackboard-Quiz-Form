using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.ComponentModel;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


namespace Blackboard_Quiz_Form
{

    public class Question
    {
        public Word.ContentControl QuestionItem { get; set; }
        public int QuestionNumber { get; set; }
        public float QuestionPosition { get; set; }
    }
    public partial class ThisDocument
    {
        private List<Question> questionList = new List<Question>();
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            Question newQuestion = new Question();
            newQuestion.QuestionItem = SelectContentControlsByTag("question")[1];
            newQuestion.QuestionNumber = 1;
            questionList.Add(newQuestion);
        }
        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }
        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.ContentControlAfterAdd += new Microsoft.Office.Interop.Word.DocumentEvents2_ContentControlAfterAddEventHandler(this.ThisDocument_ContentControlAfterAdd);
            this.ContentControlBeforeDelete += new Microsoft.Office.Interop.Word.DocumentEvents2_ContentControlBeforeDeleteEventHandler(this.ThisDocument_ContentControlBeforeDelete);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion
        private void ThisDocument_ContentControlAfterAdd(Word.ContentControl NewContentControl, bool InUndoRedo)
        {
            if (NewContentControl.Tag == "question" && InUndoRedo == false)
            {
                Question NewQuestion = new Question();
                NewQuestion.QuestionItem = NewContentControl;
                NewQuestion.QuestionItem.Title = "empty";
                List<Word.ContentControl> questionContentControls = new List<Word.ContentControl>();
                List<Question> updatedQuestionList = new List<Question>();
                for (var i = 1; i <= SelectContentControlsByTag("question").Count; i++)
                {
                    questionContentControls.Add(SelectContentControlsByTag("question")[i]);
                }
                questionContentControls = questionContentControls.FindAll(s => s.Title.Contains("Question"));
                foreach (Word.ContentControl questionContentControl in questionContentControls)
                {
                    string questionTitle = questionContentControl.Title;
                    string resultString = Regex.Match(questionTitle, @"\d+").Value;
                    int qNo = Int32.Parse(resultString);
                    updatedQuestionList.Add(questionList[qNo - 1]);
                }
                questionList = updatedQuestionList;
                questionList.Add(NewQuestion);
                NewQuestion.QuestionPosition = NewQuestion.QuestionItem.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage] + NewQuestion.QuestionItem.Range.Information[Word.WdInformation.wdActiveEndPageNumber] * 1000;
                questionList.Select(c => { c.QuestionPosition = c.QuestionItem.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage] + c.QuestionItem.Range.Information[Word.WdInformation.wdActiveEndPageNumber] * 1000; return c; }).ToList();
                questionList = questionList.OrderBy(o => o.QuestionPosition).ToList();
                questionList.Select(c => { c.QuestionNumber = questionList.IndexOf(c) + 1; return c; }).ToList();
                questionList.Select(c => { c.QuestionItem.Title = "Question " + c.QuestionNumber; return c; }).ToList();
            }
           
        }
        private void ThisDocument_ContentControlBeforeDelete(Word.ContentControl OldContentControl, bool InUndoRedo)
        {
            if (OldContentControl.Tag == "question" && InUndoRedo == false)
            {
                string questionTitle = OldContentControl.Title;
                string resultString = Regex.Match(questionTitle, @"\d+").Value;
                int qNo = Int32.Parse(resultString);
                questionList[qNo - 1].QuestionItem = null;
                questionList[qNo - 1] = null;
                questionList.RemoveAt(qNo - 1);
                Controls.Remove(OldContentControl);
                questionList.Select(c => { c.QuestionPosition = c.QuestionItem.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage] + c.QuestionItem.Range.Information[Word.WdInformation.wdActiveEndPageNumber] * 1000; return c; }).ToList();
                questionList = questionList.OrderBy(o => o.QuestionPosition).ToList();
                questionList.Select(c => { c.QuestionNumber = questionList.IndexOf(c) + 1; return c; }).ToList();
                questionList.Select(c => { c.QuestionItem.Title = "Question " + c.QuestionNumber; return c; }).ToList();
            }
        }
    }
}
