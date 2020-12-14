using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace Exams_2
{
    class Task
    {
        List<Data> TaskVariant = new List<Data>();
        public void CreateTask(ComboBox comboBox, int valueTask, int valueAnswer) {
            
            for (int i = 1; i <= valueTask; i++)
            {
                TaskVariant.Add(new Data() {VariantTask = "Task " + i, Answers = new List<string>(), CorrectAnswer = new List<bool>()});
                for (int j = 0; j < valueAnswer; j++)
                {
                    TaskVariant[i - 1].Answers.Add("");
                    TaskVariant[i - 1].CorrectAnswer.Add(false);
                }
                comboBox.Items.Add("Task " + i);
            }
            comboBox.SelectedIndex = 0;
        }
        public Data Get(int index)
        {
            if (index < 0 || index >= TaskVariant.Count)
                throw new IndexOutOfRangeException();
            return TaskVariant[index];
        }

        public void CreateDocumentWord(ProgressBar progressBar1, string[] fileName, int countVariant, int countTaskInVatiant, int Task, int numColumn, List<CheckBox> checkBoxes, List<TextBox> answertextBoxes, bool fileTask, bool fileAnswer)
        {

            DocxThread docxThread = new DocxThread(progressBar1, fileName, countVariant, countTaskInVatiant, Task, numColumn, checkBoxes, TaskVariant, answertextBoxes, fileTask, fileAnswer);

            Thread myThread = new Thread(new ParameterizedThreadStart(docxThread.ThreadDocx));
            myThread.Start(docxThread);





            //docx.PrintToWord(progressBar1, fileName, countVariant, countTaskInVatiant, Task, numColumn, checkBoxes, TaskVariant, answertextBoxes, fileTask, fileAnswer);
        }
    }

    public class DocxThread
    {
        private ProgressBar progressBar1;
        private string[] fileName;
        private int countVariant;
        private int countTaskInVariant;
        private int countTask;
        private int numColumn;
        private List<CheckBox> checkBoxes;
        private List<Data> datas;
        private List<TextBox> answertextBoxes;
        private bool fileTask;
        private bool fileAnswer;
        private Form form;

        public DocxThread(ProgressBar progressBar1, string[] fileName, int countVariant, int countTaskInVariant, int countTask, int numColumn, List<System.Windows.Forms.CheckBox> checkBoxes, List<Data> datas, List<TextBox> answertextBoxes, bool fileTask, bool fileAnswer)
        {
            this.progressBar1 = progressBar1;
            this.fileName = fileName;
            this.countVariant = countVariant;
            this.countTaskInVariant = countTaskInVariant;
            this.countTask = countTask;
            this.numColumn = numColumn;
            this.checkBoxes = checkBoxes;
            this.datas = datas;
            this.answertextBoxes = answertextBoxes;
            this.fileTask = fileTask;
            this.fileAnswer = fileAnswer;
        }

        public void ThreadDocx(object paramDocx)
        {
            Docx docx = new Docx();
            DocxThread docxThread = (DocxThread)paramDocx;
            docx.PrintToWord(docxThread.progressBar1, docxThread.fileName, docxThread.countVariant, docxThread.countTaskInVariant, docxThread.countTask, docxThread.numColumn, docxThread.checkBoxes, docxThread.datas, docxThread.answertextBoxes, docxThread.fileTask, docxThread.fileAnswer);
        }
    }
}
