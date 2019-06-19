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
        public float QuestionPosition { get; set; }
        public int QuestionNumber { get; set; }
    }
    public partial class ThisDocument
    {
        private List<Word.ContentControl> Questions = new List<Word.ContentControl>();
        int j = 1;
        private List<Question> questionList = new List<Question>();
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            Question newQuestion = new Question();
            newQuestion.QuestionItem = SelectContentControlsByTag("question")[1];
            newQuestion.QuestionPosition = SelectContentControlsByTag("question")[1].Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
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
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion

        private void ThisDocument_ContentControlAfterAdd(Word.ContentControl NewContentControl, bool InUndoRedo)
        {

            Debug.WriteLine(NewContentControl.Tag);
            if (NewContentControl.Tag == "question")
            {
                Question NewQuestion = new Question();
                questionList.Add(NewQuestion);
                NewQuestion.QuestionItem = NewContentControl;
                NewQuestion.QuestionPosition = NewContentControl.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
                if (questionList.Count >= 1)
                {
                    NewQuestion.QuestionNumber = questionList.IndexOf(NewQuestion) + 1;
                    Debug.Print(NewQuestion.QuestionNumber.ToString());
                }
                NewQuestion.QuestionItem.Title = "Question " + NewQuestion.QuestionNumber;
            }
            foreach(Question q in questionList)
            {
                //Debug.WriteLine(q.QuestionNumber);
                float position = q.QuestionItem.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
                Debug.WriteLine(q.QuestionNumber + ", " + position);
            }
            IOrderedEnumerable<Question> ordered = questionList.OrderBy(Question => Question.QuestionItem.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);
            List<Question> orderedQuestions = ordered.ToList();
            foreach (Question q in orderedQuestions)
            {
                q.QuestionNumber = orderedQuestions.IndexOf(q) + 1;
                q.QuestionItem.Title = "Question " + q.QuestionNumber;
                float position = q.QuestionItem.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
                Debug.WriteLine(q.QuestionNumber + ", " + position);
            }
        }

        private void plainTextContentControl1_Exiting(object sender, ContentControlExitingEventArgs e)
        {
            Question NewQuestion = new Question();
            NewQuestion.QuestionItem = SelectContentControlsByTag("question")[j];
            NewQuestion.QuestionPosition = SelectContentControlsByTag("question")[j].Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
            questionList.Add(NewQuestion);

            var ordered = questionList.OrderBy(Question => Question.QuestionPosition);
            questionList = ordered.ToList();
            for(int i = 1; i < questionList.Count; i++)
            {
                ordered = questionList.OrderBy(Question => Question.QuestionPosition);
                questionList = ordered.ToList();
                questionList[i].QuestionNumber = questionList.IndexOf(questionList[i]);
                questionList[i].QuestionItem.Title = "Question " + questionList[i].QuestionNumber;
                /*
                if (questionList[i].QuestionItem.Title == questionList[i - 1].QuestionItem.Title)
                {
                    questionList[i].QuestionNumber = questionList[i].QuestionNumber + 1;
                    questionList[i].QuestionItem.Title = "Question " + questionList[i].QuestionNumber;
                }
                */
            }
            j++;
            /*
            Question lastQuestion = questionList[questionList.Count - 1];
            foreach(Question q in questionList)
            {
                if(q.QuestionPosition > lastQuestion.QuestionPosition)
                {
                    lastQuestion = q;
                    lastQuestion.QuestionItem.Title = "Question " + (questionList.Count - 1);
                }
            }
            */




            /*
            foreach (Microsoft.Office.Interop.Word.ContentControl contentcontrol in this.Content.ContentControls)
            {
                
                if (contentcontrol.Tag == "repeater" && contentcontrol.Title == null)
                {
                    
                    contentcontrol.Title = "Question " + j;
                }
                
            }
            */

            /*
            for (int i = 1; i <= SelectContentControlsByTag("repeater").Count; i++)
            {

                float questionPosition = SelectContentControlsByTag("repeater")[i].Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
                Debug.WriteLine(questionPosition);
                if(i-1 != 0 && i+1 <= SelectContentControlsByTag("repeater").Count)
                { 
                    if(SelectContentControlsByTag("repeater")[i].Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage] > SelectContentControlsByTag("repeater")[i-1].Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage])
                    {
                        SelectContentControlsByTag("repeater")[i].Title = "Question " + i;
                    }
                    if (SelectContentControlsByTag("repeater")[i].Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage] < SelectContentControlsByTag("repeater")[i + 1].Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage])
                    {
                        SelectContentControlsByTag("repeater")[i].Title = "Question " + (i - 1);
                    }
                }
                
            }
            */
        }
        /*
        private void plainTextContentControl3_Exiting(object sender, ContentControlExitingEventArgs e)
        {
            Question NewQuestion = new Question();
            NewQuestion.QuestionItem = SelectContentControlsByTag("question")[j];
            NewQuestion.QuestionPosition = SelectContentControlsByTag("question")[j].Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
            questionList.Add(NewQuestion);

            var ordered = questionList.OrderBy(Question => Question.QuestionPosition);
            questionList = ordered.ToList();
            for (int i = 1; i < questionList.Count; i++)
            {
                ordered = questionList.OrderBy(Question => Question.QuestionPosition);
                questionList = ordered.ToList();
                questionList[i].QuestionNumber = questionList.IndexOf(questionList[i]);
                questionList[i].QuestionItem.Title = "Question " + questionList[i].QuestionNumber;

                if (questionList[i].QuestionItem.Title == questionList[i - 1].QuestionItem.Title)
                {
                    questionList[i].QuestionNumber = questionList[i].QuestionNumber + 1;
                    questionList[i].QuestionItem.Title = "Question " + questionList[i].QuestionNumber;
                }
                if (questionList[i].QuestionNumber < questionList[i - 1].QuestionNumber)
                {
                    questionList[i].QuestionNumber = questionList[i - 1].QuestionNumber + 1;
                    questionList[i].QuestionItem.Title = "Question " + questionList[i].QuestionNumber;
                }
            }
            j++;
        }
        */
    }
}
