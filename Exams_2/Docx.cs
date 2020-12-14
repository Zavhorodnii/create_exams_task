using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;

namespace Exams_2
{
    class Docx
    {
        ProgressBar progressBar;
        object oMissing = Missing.Value;
        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

        _Application oWord;
        Word.Document oDoc, oDocHeader, oDocFooter;
        private List<int> randomTask = new List<int>();
        private List<int> temporaryRandomTask = new List<int>();
        private int randomTaskIndex = 0;

        public void PrintToWord(ProgressBar progressBar1, string[] fileName, int countVariant, int countTaskInVariant, int countTask, int numColumn, List<System.Windows.Forms.CheckBox> checkBoxes, List<Data> datas, List<TextBox> answertextBoxes, bool fileTask, bool fileAnswer)
        {
            progressBar = progressBar1;
            randomTaskIndex = 0;
            int randomNum = 0;
            temporaryRandomTask.Clear();
            randomTask.Clear();

            int countRepeat = 1;

            countRepeat = countVariant * countTaskInVariant / countTask;
            
            Random random = new Random();
            //Debug.WriteLine("countRepeat +" + countRepeat + " countTask +" + countTask + " countTaskInVariant +" + countTaskInVariant);
            if (countVariant * countTaskInVariant > countTask)
            {
                randomTask.Clear();
                //Debug.WriteLine("countRepeat * countTas =" + countRepeat * countTask);
                for (int j = 0; j < countRepeat * countTask;)
                {
                    randomNum = random.Next(0, countTask);

                    if (CheckRandomNum(randomNum, countRepeat, randomTask))
                    {
                        randomTask.Add(randomNum);
                        j++;
                    }
                }
                temporaryRandomTask.Clear();

                //Debug.WriteLine("i =" + (countVariant * countTaskInVariant - countTask * countRepeat));
                for (int i = randomTask.Count - countTask/2; i < randomTask.Count; i++)
                    temporaryRandomTask.Add(randomTask[i]);
                /*foreach (var item in temporaryRandomTask)
                {
                    Debug.Write(item + " ");
                }
                Debug.WriteLine("\n" + temporaryRandomTask.Count);
                Debug.WriteLine("2");*/
                for (int j = 0; j < countVariant * countTaskInVariant - countTask * countRepeat;)
                {
                    randomNum = random.Next(0, countTask); 
                    if (CheckRandomNum(randomNum, 1, temporaryRandomTask))
                    {
                        temporaryRandomTask.Add(randomNum);
                        randomTask.Add(randomNum);
                        j++;
                    }
                }
                //Debug.WriteLine("3");
            }
            else
            {
                for (int j = 0; j < countTaskInVariant * countVariant;)
                {
                    randomNum = random.Next(0, countTask);
                    if (CheckRandomNum(randomNum, 1, randomTask))
                    {
                        randomTask.Add(randomNum);
                        j++;
                    }
                }
            }
            /*
            Debug.WriteLine("hello");

            foreach (var item in randomTask)
            {
                Debug.Write(item + " ");
            }
            Debug.WriteLine("\n" + randomTask.Count);*/
            try
            {
                if(fileTask)
                    CreateTaskdDocument(fileName[0], fileName[1], countVariant, countTaskInVariant, datas, answertextBoxes);
                if(fileAnswer)
                    CreateAnsverDocument(fileName[2], fileName[3], countVariant, countTaskInVariant, numColumn, datas, answertextBoxes);
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }
        

        private bool CheckRandomNum(int randomNum, int countRepeat, List<int> listRandomTask)
        {
            int num = 0;
            if (listRandomTask.Count == 0)
                return true;
            for (int i = 0; i < listRandomTask.Count; i++)
            {

                //Debug.WriteLine("randomNum =" + randomNum + " listRandomTask[i] =" + listRandomTask[i]);
                if (randomNum == listRandomTask[i])
                {
                    num++;
                    if (num >= countRepeat)
                        return false;
                }
            }
            return true;
        }

        private void CreateTaskdDocument(string headFileName, string footFileName, int countVariant, int countTaskInVariant, List<Data> datas, List<TextBox> answertextBoxes)
        {
            oWord = new Word.Application();
            oWord.Visible = true;
            if (headFileName != null)
                oDocHeader = oWord.Documents.Open(headFileName);
            if (footFileName != null)
                oDocFooter = oWord.Documents.Open(footFileName);

            oDoc = oWord.Documents.Add();

            for (int i = 0; i < countVariant; i++)
            {
                if (headFileName != null)
                    PasteHeader();
                string[] str;
                for (int j = 0; j < countTaskInVariant; j++)
                {
                    if (datas[randomTask[randomTaskIndex]].Task != null)
                        str = datas[randomTask[randomTaskIndex]].Task.Split(new String[] { "\r\n" }, StringSplitOptions.None);
                    else
                        str = new string[] { "" };
                    for (int k = 0; k < str.Count(); k++)
                    {
                        if (k == 0)
                            CreateTask(str[k], (j + 1).ToString() + ".");
                        else
                            CreateTask(str[k], "0");
                    }
                    for (int k = 0; k < datas[randomTask[randomTaskIndex]].Answers.Count; k++)
                    {
                        if (datas[randomTask[randomTaskIndex]].Answers[k] != null)
                            str = datas[randomTask[randomTaskIndex]].Answers[k].Split(new String[] { "\r\n" }, StringSplitOptions.None);
                        else
                            str = new string[] { "" };
                        for (int l = 0; l < str.Count(); l++)
                        {
                            if (l == 0)
                                CreateAnswerVariant(str[l], answertextBoxes[k].Text);
                            else
                                CreateAnswerVariant(str[l], "-1");
                        }
                        UpdateProgressBar();
                    }
                    randomTaskIndex++;
                }
                if (footFileName != null)
                    PasteFooter();

                BreakPage();
            }
        }

        private void PasteHeader()
        {
            oDocHeader.Range().Copy();
            oDoc.Range(oDoc.Range().End - 1).PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);
        }

        private void PasteFooter()
        {
            oDocFooter.Range().Copy();
            oDoc.Range(oDoc.Range().End - 1).Paste();
        }

        private void CreateTask(string str, string variantAnswer)
        {
            Word.Paragraph paragraph;
            paragraph = oDoc.Content.Paragraphs.Add(ref oMissing);
            paragraph.TabHangingIndent(1);
            if (!variantAnswer.Equals("0"))
                paragraph.Range.Text = variantAnswer + "\t" + str;
            else
                paragraph.Range.Text = str;
            paragraph.Format.FirstLineIndent = oWord.CentimetersToPoints(-0.6f);
            paragraph.Format.LeftIndent = oWord.CentimetersToPoints(1.25f);
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Format.SpaceAfter = 0;
            paragraph.LineSpacing = 12F;
            paragraph.Range.Paragraphs.LineSpacing = 12F;
            paragraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify;
            paragraph.Range.InsertParagraphAfter();
        }

        private void CreateAnswerVariant(string str, string variantAnswer)
        {
            Word.Paragraph paragraph;
            paragraph = oDoc.Content.Paragraphs.Add(ref oMissing);
            paragraph.TabHangingIndent(1);
            if (!variantAnswer.Equals("-1"))
                paragraph.Range.Text = (variantAnswer + "\t" + str).TrimStart( new char[] { '\t' });
            else
                paragraph.Range.Text = str;
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Format.SpaceAfter = 0;
            paragraph.LineSpacing = 12F;
            paragraph.Range.Paragraphs.LineSpacing = 12F;
            paragraph.Format.FirstLineIndent = oWord.CentimetersToPoints(-0.6f);
            paragraph.Format.LeftIndent = oWord.CentimetersToPoints(2.57f);
            paragraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify;
            paragraph.Range.InsertParagraphAfter();
        }

        private void BreakPage()
        {
            Word.Paragraph paragraph;
            paragraph = oDoc.Content.Paragraphs.Add(ref oMissing);
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = Word.WdBreakType.wdPageBreak;
            paragraph.Range.Collapse(ref oCollapseEnd);
            paragraph.Range.InsertBreak(ref oPageBreak);
            paragraph.Range.Collapse(ref oCollapseEnd);
        }

        private void CreateAnsverDocument(string headFileName, string footFileName, int countVariant, int countTaskInVariant, int numColumn, List<Data> datas, List<TextBox> answertextBoxes)
        {
            oWord = new Word.Application();
            oWord.Visible = true;

            if(headFileName != null)
                oDocHeader = oWord.Documents.Open(headFileName);
            if(footFileName != null)
                oDocFooter = oWord.Documents.Open(footFileName);

            oDoc = oWord.Documents.Add();


            randomTaskIndex = 0;
            for (int i = 0; i < countVariant; i++)
            {
                if(headFileName != null)
                    PasteHeader();
                CreateAnswers(countTaskInVariant, numColumn, datas, answertextBoxes);
                if(footFileName != null)
                    PasteFooter();
                BreakPage();
            }
        }

        private void CreateAnswers(int countTaskInVariant, int numColumn, List<Data> datas, List<TextBox> answertextBoxes)
        {
            int countVariaantAnsverInTask = datas[randomTask[randomTaskIndex]].Answers.Count +1;
            //Debug.WriteLine("countVariaantAnsverInTask =" + countVariaantAnsverInTask);
            int rowTable = (int)Math.Ceiling((decimal)countTaskInVariant / numColumn) + 1;

            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, rowTable, countVariaantAnsverInTask * numColumn, ref oMissing, ref oMissing);

            oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

            oTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            oTable.Range.Font.Name = "Times New Roman";
            oTable.Range.Font.Size = 13;

            oTable.Range.Paragraphs.LineSpacing = 12F;
            oTable.Range.Paragraphs.SpaceAfter = 0;

            bool delete = false;
            int columnFrom = 1, columnTo = countVariaantAnsverInTask, rowNum = 1, indexColumName = 0, indexCheckAnsver = 0, indexDatas = 0;
            for (int wall = 0; wall < numColumn; wall++)
            {
                for (int row = 1; row <= rowTable; row++)
                {
                    if (rowNum == countTaskInVariant + 1)
                    {
                        oDoc.Range(oTable.Cell(row, columnFrom).Range.Start, oTable.Cell(row, columnTo).Range.End).Cells.Delete();
                        delete = true;
                        continue;
                    }
                    for (indexCheckAnsver = 0; columnFrom <= columnTo; columnFrom++)
                    {
                        if (row == 1)
                        {
                            if (columnFrom == columnTo - countVariaantAnsverInTask + 1)
                            {
                                oTable.Cell(row, columnFrom).Range.Text = "№";
                                oTable.Cell(row, columnFrom).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray20;
                            }
                            else
                            {
                                oTable.Cell(row, columnFrom).Range.Text = answertextBoxes[indexColumName++].Text[0].ToString();
                                oTable.Cell(row, columnFrom).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray20;
                            }
                        }
                        else
                        {
                            if (columnFrom == columnTo - countVariaantAnsverInTask + 1)
                            {
                                oTable.Cell(row, columnFrom).Range.Text = rowNum.ToString();
                                oTable.Cell(row, columnFrom).Range.Shading.BackgroundPatternColor = WdColor.wdColorGray20;
                                rowNum++;
                            }
                            else
                            {
                                if (datas[randomTask[indexDatas]].CorrectAnswer[indexCheckAnsver])
                                {
                                    oTable.Cell(row, columnFrom).Range.Shading.BackgroundPatternColor = WdColor.wdColorBlack;
                                }
                                indexCheckAnsver++;
                                UpdateProgressBar();
                            }
                        }
                    }
                    if (row > 1)
                        indexDatas++;
                    if (row < rowTable && !delete)
                        columnFrom -= countVariaantAnsverInTask;
                }
                if(!delete)
                columnTo += countVariaantAnsverInTask;
                indexColumName = 0;
            }
        }

        private void UpdateProgressBar()
        {
            //Debug.WriteLine("------------- " + progressBar.Value);
            
            progressBar.BeginInvoke(new InvokeDelegate(InvokeMethod));
            
            //progressBar.PerformStep();
        }

        public delegate void InvokeDelegate();
        
        public void InvokeMethod()
        {
            progressBar.PerformStep();
        }
    }
}