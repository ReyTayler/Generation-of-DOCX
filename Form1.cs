using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace WordChanger
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = @"All Files | *.*";
            openFileDialog.FileName = "";
            openFileDialog.InitialDirectory = @"c:\\Users\\";
            openFileDialog.RestoreDirectory = true;

            string fileName = "";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
                fileName = System.IO.Path.GetFullPath(openFileDialog.FileName);

            WordChange.ConvertToPDF(fileName);

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            // создаём экземпляр changer класса WordChange и закидывем в него имя шаблона документа, с которым будем работать

            WordChange changer = new WordChange("Шаблон документа для кафедры АСОИУ.docx");

            // создаём словарь items
            var items = new Dictionary<string, string>()
            {
                {"{where}", textBox2.Text},
                {"{from whom}", textBox3.Text},
                {"{headline}",richTextBox3.Text},
                {"{text}", richTextBox1.Text},
                {"{date}", dateTimePicker1.Value.ToString("dd.MM.yyyy")},
            };

            changer.GenerateDocX(items);
        }
    }
}
