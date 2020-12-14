using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Exams_2
{
    class UpdateAnswerPanel
    {
        private Panel panel;
        private Button button;
        private int valueColumn = 0;
        private List<PanelsClass> panelsClass = new List<PanelsClass>();
        

        public class PanelsClass
        {
            public Panel panel;
            public Label label;
            public TextBox textBox;
            public PanelsClass()
            {
                label = new Label();
                textBox = new TextBox();
                panel = new Panel(); 
            }
        }

        public List<PanelsClass> GetList()
        {
            return panelsClass;
        }

        public UpdateAnswerPanel(Panel panel, Button button)
        {
            this.button = button;
            this.panel = panel;
        }


        public void ChangeAnswerBox(int chngePanel, Label label13)
        {
            if(chngePanel > valueColumn)
            {
                for (int i = valueColumn; i < chngePanel; i++)
                {
                    PanelsClass pc = new PanelsClass();
                    pc.label.TextAlign = ContentAlignment.MiddleCenter;
                    pc.label.Text = (valueColumn + 1).ToString();
                    pc.label.Height = 26;
                    pc.label.Width = 26;
                    pc.label.Font = new Font("Times New Roman", 12);
                    pc.label.Location = new Point(0, 0);
                    
                    pc.textBox.Name = "ColumnTextBox" + valueColumn;
                    pc.textBox.Font = new Font("Times New Roman", 12);
                    pc.textBox.TextAlign = HorizontalAlignment.Center;
                    pc.textBox.TextChanged += WriteText;
                    pc.textBox.Height = 26;
                    pc.textBox.Width = 26;
                    pc.textBox.Location = new Point(26, 0);

                    pc.panel.Width = 52;
                    pc.panel.Height = 26;
                    pc.panel.Location = new Point(panelsClass.Count * 52, 0);
                    pc.panel.Controls.Add(pc.textBox);
                    pc.panel.Controls.Add(pc.label);
                    panelsClass.Add(pc);

                    this.panel.AutoScrollPosition = new Point(0, 0);
                    this.panel.Controls.Add(pc.panel);
                    valueColumn++;
                    label13.Text = valueColumn.ToString();
                }
            }
            else
            {
                for (; valueColumn > chngePanel;)
                {
                    this.panel.Controls.Remove(panelsClass[panelsClass.Count - 1].panel);
                    panelsClass.RemoveAt(panelsClass.Count - 1);
                    valueColumn--;
                    label13.Text = valueColumn.ToString();
                }
            }
        }

        private void WriteText(object sender, EventArgs e)
        {
            for(int i = 0; i < panelsClass.Count(); i++)
            {
                if ((panelsClass[i].textBox.Text = panelsClass[i].textBox.Text.TrimStart( new char[] {' '})) == "")
                {
                    button.Enabled = false;
                    return;
                }
            }

            button.Enabled = true;
        }
    }
}
