using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace Exams_2
{
    public partial class Form1 : Form
    {
        Task task = new Task();
        List<TextBox> textBoxes = new List<TextBox>();
        List<TextBox> answertextBoxes = new List<TextBox>();
        List<CheckBox> checkBoxes = new List<CheckBox>();
        string[] fileName = new string[4];
        UpdateAnswerPanel updateAnswerPanel;

        public Form1()
        {
            InitializeComponent();
            openFileDialog1.Filter = "All files(*.*)|*.*";
            saveFileDialog1.Filter = "All files(*.*)|*.*";
            updateAnswerPanel = new UpdateAnswerPanel(panel3, button3);
            updateAnswerPanel.ChangeAnswerBox(Convert.ToInt32(numericUpDown2.Value), label13);
            //Change();
        }

        private void CreateTask(object swnder, EventArgs e)
        {
            button1.Enabled = true;
            button3.Enabled = false;
            button5.Enabled = true;
            panel3.Enabled = false;
            textBox1.Enabled = true;
            numericUpDown4.Enabled = false;
            numericUpDown2.Enabled = false;
            CreatePanelAnswers();
            task.CreateTask(comboBox1, (int)numericUpDown4.Value, updateAnswerPanel.GetList().Count);
        }

        private void PrintToWord(object sender, EventArgs e)
        {
            //Debug.WriteLine("updateAnswerPanel.GetList().Count =" + updateAnswerPanel.GetList().Count);
            for (int i = 0; i < updateAnswerPanel.GetList().Count; i++)
            {
                answertextBoxes.Add(updateAnswerPanel.GetList()[i].textBox);
            }

            progressBar1.Value = 0;
            if (checkBox1.Checked && checkBox2.Checked)
                progressBar1.Step = (int)((100 / (numericUpDown3.Value * numericUpDown6.Value * updateAnswerPanel.GetList().Count) / 2)*1000);
            else
                progressBar1.Step = (int)((100 / (numericUpDown3.Value * numericUpDown6.Value * updateAnswerPanel.GetList().Count))*1000);

            task.CreateDocumentWord(progressBar1, fileName, (int)numericUpDown3.Value, (int)numericUpDown6.Value, (int)numericUpDown4.Value, (int)numericUpDown1.Value, checkBoxes, answertextBoxes, checkBox1.Checked, checkBox2.Checked);
            
        }

        private void CreatePanelAnswers()
        {
            for (int i = 0, k = 0; k < updateAnswerPanel.GetList().Count; k++)
            {
                CheckBox checkBox = new CheckBox();
                ContentAlignment contentAlignment = checkBox.TextAlign = ContentAlignment.MiddleCenter;
                //checkBox.Text = answertextBoxes[k - 1].Text;
                checkBox.Text = updateAnswerPanel.GetList()[k].textBox.Text;
                checkBox.Height = 41;
                checkBox.Width = 50;
                checkBox.Font = new Font("Times New Roman", 12);
                checkBox.Location = new Point(0, i);
                checkBoxes.Add(checkBox);
                checkBox.CheckedChanged += textBoxAnswers_checked;
                TextBox textBox = new TextBox();
                textBox.Multiline = true;
                textBox.ScrollBars = ScrollBars.Vertical;
                textBox.Width = panel1.Size.Width - 71;
                textBox.Height = 40;
                textBoxes.Add(textBox);
                textBox.TextChanged += textBoxAnswers_TextChanged;
                textBox.Font = new Font("Times New Roman", 11);
                textBox.Location = new Point(50, i);

                panel1.AutoScrollPosition = new Point(0, 0);
                panel1.Controls.Add(textBox);
                panel1.Controls.Add(checkBox);
                i += 41;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
                return;
            Data data = task.Get(comboBox1.SelectedIndex);
            label6.Text = data.VariantTask;
            textBox1.Text = data.Task;
            //Debug.WriteLine("textBoxes.Count =" + textBoxes.Count + " data.Answers =" + data.Answers.Count);
            for(int i = 0; i < textBoxes.Count; i++ )
            {
                textBoxes[i].Text = data.Answers[i];
                checkBoxes[i].Checked = data.CorrectAnswer[i];
            }
            if (comboBox1.SelectedIndex == 0)
                button4.Enabled = false;
            else
                button4.Enabled = true;
            if (comboBox1.SelectedIndex == comboBox1.Items.Count - 1)
                button5.Enabled = false;
            else
                button5.Enabled = true;
        }
        

        ///----------

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Data data = task.Get(comboBox1.SelectedIndex);
            data.Task = textBox1.Text;
        }

        private void textBoxAnswers_TextChanged(object sender, EventArgs e)
        {
            Data data = task.Get(comboBox1.SelectedIndex);
            data.Answers[textBoxes.IndexOf((TextBox)sender)] = ((TextBox)sender).Text;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if(comboBox1.SelectedIndex < comboBox1.Items.Count - 1)
                comboBox1.SelectedIndex++;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > 0)
                comboBox1.SelectedIndex--;
        }


        private void textBoxAnswers_checked(object sender, EventArgs e)
        {
            Data data = task.Get(comboBox1.SelectedIndex);
            data.CorrectAnswer[checkBoxes.IndexOf((CheckBox)sender)] = ((CheckBox)sender).Checked;
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

            updateAnswerPanel.ChangeAnswerBox(Convert.ToInt32(numericUpDown2.Value), label13);
            
        }

        private void TaskHeaderfile(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            fileName[0] = openFileDialog1.FileName;
        }

        private void TaskFooterFile(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            fileName[1] = openFileDialog1.FileName;
        }

        private void AnswerHeaderFile(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            fileName[2] = openFileDialog1.FileName;
        }

        private void AnswerFooterFile(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            fileName[3] = openFileDialog1.FileName;
        }

    }
}
