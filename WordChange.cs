using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Document.NET;
using DocX = Xceed.Words.NET.DocX;
using Microsoft.Office.Interop.Word;
using Aspose.Words;

namespace WordChanger
{
    class WordChange: MainForm
    {
        public FileInfo fileInfo;
        DocX doc;
        // конструктор класса WordChange
        //
        // входные параметры:
        //      string filename - имя docx-файла c шаблоном 
        //
        // сводка:
        //      проверяет существование файла в директории и присваивает
        //      переменной типа FileInfo fileInfo экземпляр класса FileInfo,
        //      который хранит полную информацию о шаблоне.
        public WordChange(string fileName)
        {
            // делаем проверку на существование файла шаблона на компьютере

            if (File.Exists(fileName))
            {
                // Создаём экземпляр класса FileInfo, который будет работать c нашим шаблоном fileName

                fileInfo = new FileInfo(fileName);

                //загрузим шаблон файла, хранящийся на ПК

                doc = DocX.Load(fileInfo.FullName);
            }

            else
                throw new FileNotFoundException("Файл не найден!");
        }

        internal static void ConvertToPDF(string fileName)
        {
            Microsoft.Office.Interop.Word.Application pdf = new Microsoft.Office.Interop.Word.Application();
            object missing = System.Reflection.Missing.Value;

            FileInfo wordFile = new FileInfo(fileName);
            object nameFile = wordFile.FullName;

            Microsoft.Office.Interop.Word.Document document = pdf.Documents.Open(ref nameFile, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            document.Activate();

            object outputFileName = wordFile.FullName.Replace(".docx", ".pdf");
            object fileFormat = WdSaveFormat.wdFormatPDF;

            document.SaveAs2(ref outputFileName, ref fileFormat, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            object saveChanges = WdSaveOptions.wdSaveChanges;
            document.Close(ref saveChanges, ref missing, ref missing);
            document = null;

            pdf.Quit(ref missing, ref missing, ref missing);
            pdf = null;

            MessageBox.Show("Данный файл был преобразован в PDF!");
        }

        internal void SaveNewDocX()
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "docx files (*.docx)|*.docx|All files (*.*)|*.*";
            saveFile.Title = "Сохранить сгенерированный документ в DocX";
            saveFile.RestoreDirectory = true;
            saveFile.FileName = DateTime.Now.ToString("dd_MM_yyyy HH_mm_ss") + " ASPAA_KAI";
            saveFile.InitialDirectory = @"c:\\Users\\";

            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                doc.SaveAs(saveFile.FileName);
            }
        }

        [Obsolete]
        internal void GenerateDocX(Dictionary<string, string> items)
        {
            try
            {
                foreach (var item in items)
                {
                    doc.ReplaceText(item.Key, item.Value);
                }
                MessageBox.Show("Ваш документ успешно сгенерирован!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                SaveNewDocX();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
