using ExamProject;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace DocumentProcessing
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            var helper = new DocumentHelper("шаблон.docx");
            var items = new Dictionary<string, string>
            {
                {"<Число>", textBox1.Text},
                {"<Месяц и год>", textBox2.Text},
                {"<Номер отчёта>", textBox3.Text},
                {"<Наименование>", textBox4.Text}
            };
            helper.Process(items);
        }
    }
}
