using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Table
{
    internal static class WritingHelper
    {
        public static void ChoseSortMethod()
        {
            Console.WriteLine($"\nВыберите тип сортировки:\n" +
                $"1 - По общему количеству документов (в случае равенства – по количеству РКК).\n" +
                $"2 - По количеству обращений (в случае равенства – по количеству РКК).\n" +
                $"3 - По количеству РКК (в случае равенства – по количеству обращений).\n" +
                $"4 - По фамилии ответственного исполнителя.\n" +
                $"Укажите цифру от 1 до 4");
        }

        public static void TableInfo(Data data)
        {
            Console.WriteLine(WordInfo(data));
            TableBorder();
            Console.WriteLine(String.Format("| {0,-4} | {1,-20} | {2,-24} | {3,-28} | {4,-22} |",
                "№",
                "Ответственный",
                "Количество неисполненных",
                "Количество неисполненных",
                "Общее количество"));
            Console.WriteLine(String.Format("| {0,-4} | {1,-20} | {2,-24} | {3,-28} | {4,-22} |",
                "п.п.",
                "исполнитель",
                "входящих документов",
                "письменных обращений граждан",
                "документов и обращений"));
        }

        public static string WordInfo(Data data)
        {
            return ($"Справка о неисполненных документах и обращениях граждан\n" +
                $"Не исполнено в срок {data.CountSum} документов, из них:\n" +
                $"- количество неисполненных входящих документов: {data.CountRkk};\n" +
                $"- количество неисполненных письменных обращений граждан: {data.CountOrb}.\n" +
                $"Сортировка: {data.SortMethod}.");
        }

        public static void TableBorder()
        {
            Console.WriteLine(new String('-', 114));
        }
    }
}
