using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
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
    public class Question
    {
        public Word.ContentControl QuestionItem { get; set; }
        public int QuestionNumber { get; set; }
        public float QuestionPosition { get; set; }
    }
    public partial class ThisDocument
    {
        Question lastQuestionOnPreviousPage;
        private List<Question> questionList = new List<Question>();
        List<List<Question>> listOfQuestionListsByPage = new List<List<Question>>();
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            Question newQuestion = new Question();
            newQuestion.QuestionItem = SelectContentControlsByTag("question")[1];
            newQuestion.QuestionNumber = 1;
            questionList.Add(newQuestion);
            List<Question> newPageList = new List<Question>();
            listOfQuestionListsByPage.Add(newPageList);
            newPageList.Add(newQuestion);
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
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion
        private void ThisDocument_ContentControlAfterAdd(Word.ContentControl NewContentControl, bool InUndoRedo)
        {
            if (NewContentControl.Tag == "question")
            {
                Question NewQuestion = new Question();
                NewQuestion.QuestionItem = NewContentControl;
                questionList.Add(NewQuestion);
                int newPageCount = this.Content.get_Information(Microsoft.Office.Interop.Word.WdInformation.wdNumberOfPagesInDocument);
                foreach (Question q in questionList)
                {
                    q.QuestionPosition = q.QuestionItem.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
                }
                if (newPageCount > 1)
                {
                    List<Question> previousPageQuestions = questionList.FindAll(x => x.QuestionItem.Range.Information[Word.WdInformation.wdActiveEndPageNumber] == newPageCount - 1);
                    lastQuestionOnPreviousPage = previousPageQuestions.OrderByDescending(item => item.QuestionPosition).First();
                    Debug.WriteLine("The last question on the previous page is " + lastQuestionOnPreviousPage.QuestionNumber);
                }
                foreach (Question q in questionList)
                {
                    if (q.QuestionItem.Range.Information[Word.WdInformation.wdActiveEndPageNumber] > 1)
                    {
                        q.QuestionPosition += lastQuestionOnPreviousPage.QuestionPosition;
                    }
                }
                IOrderedEnumerable<Question> ordered = questionList.OrderBy(Question => Question.QuestionPosition);
                List<Question> orderedQuestions = ordered.ToList();
                foreach (Question q in orderedQuestions)
                {
                    if (orderedQuestions.IndexOf(q) > 0)
                    { 
                        q.QuestionNumber = orderedQuestions.IndexOf(q) + 1;
                        q.QuestionItem.Title = "Question " + q.QuestionNumber;
                    }
                    /*
                if (orderedQuestions.IndexOf(q) > 0)
                {
                    List<Question> questionSubList = orderedQuestions.GetRange(0, orderedQuestions.IndexOf(q));

                }

                q.QuestionNumber = orderedQuestions.IndexOf(q) + 1;
                q.QuestionItem.Title = "Question " + q.QuestionNumber;
                Debug.WriteLine(q.QuestionItem.Title + "; " + q.QuestionPosition);
                */
                }
            }
        }
    }
}
