using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace Table
{
    public class Program
    {
        public static void Main()
        {
            //Выбираем файлы
            Console.WriteLine("Укажите путь до файла с обращениями");
            string obrPath = Console.ReadLine();

            Console.WriteLine("Укажите путь до файла с РКК");
            string rkkPath = Console.ReadLine();

            //Выбираем метод сортировки
            WritingHelper.ChoseSortMethod();
            int sortMethod = Convert.ToInt32(Console.ReadLine());

            //Запускаем таймер
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            //Читаем файлы
            var obr = new StreamReader(obrPath).ReadToEnd();
            var rkk = new StreamReader(rkkPath).ReadToEnd();

            //Создаем словарь и сортируем его
            var data = new Data();
            var dict = data.Counting(new Dictionary<string, Data>(), obr, Docs.obr);
            dict = data.Counting(dict, rkk, Docs.rkk);
            var sortedDict = data.DictionarySort(dict, sortMethod);

            //Рисуем и заполняем таблицу
            Console.WriteLine();
            WritingHelper.TableInfo(data);
            WritingHelper.TableBorder();
            for (int i = 0; i < sortedDict.Count; i++)
            {
                Console.WriteLine(String.Format("| {0,-4} | {1,-20} | {2,-24} | {3,-28} | {4,-22} |", 
                    i + 1, 
                    sortedDict[i].Key, 
                    sortedDict[i].Value.Rkk, 
                    sortedDict[i].Value.Obr, 
                    sortedDict[i].Value.Sum));
            }
            WritingHelper.TableBorder();

            //Останавливаем таймер
            stopwatch.Stop();
            Console.WriteLine($"\nВремя выполнения алгоритма: {stopwatch.ElapsedMilliseconds} мс." );

            Console.WriteLine("\nХотите вывести результат работы программы в файл RTF?\n" +
                "1 - Да, 2 - Нет, выключить программу");

            //Создаем RTF документ
            if (Convert.ToInt32(Console.ReadLine()) == 1)
            {
                Console.WriteLine("\nСоздаем документ...");
                var ww = new WordWriter();
                ww.WriteParagraph("Справка о неисполненных документах и обращениях граждан", bold: true, fontSize: 18);
                ww.WriteParagraph(WritingHelper.WordInfo(data));
                ww.WriteTable(sortedDict);
                ww.WriteParagraph("Дата составления справки: ");
                ww.WriteDate();
                ww.Close();
                Console.WriteLine("\nФайл сохранен в папку \"Документы\"\n" +
                    "с названием \"Тестовое задание - результат работы программы\"\n");
            }

            Console.WriteLine("Для выхода из программы нажмите любую кнопку...");
            Console.ReadLine();
        }
    }
}