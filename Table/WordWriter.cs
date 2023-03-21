using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Table
{
    internal class WordWriter
    {
        public Word.Application App { get; private set; }
        public Word.Document Document { get; private set; }
        public Word.Paragraph Paragraph { get; private set; }

        public WordWriter()
        {
            App = new Word.Application();
            Document = App.Documents.Add(Visible: true);
            Paragraph = Document.Paragraphs.Add();
        }

        public void WriteParagraph(string str, bool bold = false, int fontSize = 11)
        {
            Paragraph = Document.Paragraphs.Add();
            Word.Range Range = Paragraph.Range;
            if (bold)
                Range.Bold = 1;
            Range.Font.Size = fontSize;
            Range.Text = str;
            Range.InsertParagraphAfter();            
        }

        public void WriteDate()
        {
            Paragraph = Document.Paragraphs.Add();
            Word.Range Range = Paragraph.Range;
            Range.Bold = 1;
            Range.InsertDateTime();
            Range.InsertParagraphAfter();
        }

        public void WriteTable(List<KeyValuePair<string, Data>> sortedDict)
        {
            Paragraph = Document.Paragraphs.Add();
            Word.Range Range = Paragraph.Range;

            Word.Table t = Document.Tables.Add(Range, sortedDict.Count, 5);
            t.Borders.Enable = 1;
            t.Cell(1, 1).Range.Text = "№ п.п.";
            t.Cell(1, 2).Range.Text = "Ответственный исполнитель";
            t.Cell(1, 3).Range.Text = "Количество неисполненных входящих документов";
            t.Cell(1, 4).Range.Text = "Количество неисполненных письменных обращений граждан";
            t.Cell(1, 5).Range.Text = "Общее количество документов и обращений";

            for (int i = 0; i < sortedDict.Count; i++)
            {
                t.Cell(i + 2, 1).Range.Text = (i + 1).ToString();
                t.Cell(i + 2, 2).Range.Text = sortedDict[i].Key.ToString();
                t.Cell(i + 2, 3).Range.Text = sortedDict[i].Value.Rkk.ToString();
                t.Cell(i + 2, 4).Range.Text = sortedDict[i].Value.Obr.ToString();
                t.Cell(i + 2, 5).Range.Text = sortedDict[i].Value.Sum.ToString();
            }
        }

        public void Close()
        {
            Document.SaveAs2(FileName: "Тестовое задание - результат работы программы.rtf", FileFormat: Word.WdSaveFormat.wdFormatRTF);
            Document.Close();
        }
    }
}
