using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Table
{
    enum Docs
    {
        obr,
        rkk,
    }

    internal class Data
    {
        public int Obr { get; private set; }
        public int Rkk { get; private set; }
        public int Sum { get; private set; }
        public int CountOrb { get; private set; }
        public int CountRkk { get; private set; }
        public int CountSum { get; private set; }
        public string SortMethod { get; private set; } = "По общему количеству документов";

        public Dictionary<string, Data> Counting(Dictionary<string, Data> dict, string stream, Docs doc)
        {
            string[] text = stream.Split(new string[] { "\t", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < text.Length; i += 2)
            {
                string name = GetShortName(text[i]);

                if (name == "Климов С.А.")
                    name = GetPerformer(text[i + 1]);

                if (!dict.ContainsKey(name))
                {
                    switch (doc)
                    {
                        case Docs.obr:
                            dict.Add(name, new Data { Obr = 1, Sum = 1 });
                            CountOrb++; CountSum++;
                            break;

                        case Docs.rkk:
                            dict.Add(name, new Data { Rkk = 1, Sum = 1 });
                            CountRkk++; CountSum++;
                            break;
                    }
                }
                else
                {
                    switch (doc)
                    {
                        case Docs.obr:
                            dict[name].Obr++;
                            dict[name].Sum++;
                            CountOrb++; CountSum++;
                            break;

                        case Docs.rkk:
                            dict[name].Rkk++;
                            dict[name].Sum++;
                            CountRkk++; CountSum++;
                            break;
                    }
                }
            }
            return dict;
        }

        private string GetShortName(string str)
        {
            string[] temp = str.Split();
            string handler = $" {temp[1][..1]}.{temp[2][..1]}.";
            string name = new StringBuilder(temp[0]).Append(handler).ToString();
            return name;
        }

        private string GetPerformer(string str)
        {
            int index = str.IndexOf('.') + 1;
            index = str.IndexOf('.', index) + 1;
            string performer = str[..index];
            return performer;
        }

        public List<KeyValuePair<string, Data>> DictionarySort(Dictionary<string, Data> dict, int sort)
        {
            var sorted_dict = new List<KeyValuePair<string, Data>>();
            switch (sort)
            {
                default:
                case 1:
                    sorted_dict = dict.OrderByDescending(x => x.Value.Sum).ThenBy(x => x.Value.Obr).ThenBy(x => x.Value.Rkk).ThenBy(x => x.Key).ToList();
                    SortMethod = "По общему количеству документов";
                    break;
                case 2:
                    sorted_dict = dict.OrderByDescending(x => x.Value.Obr).ThenBy(x => x.Value.Sum).ThenBy(x => x.Value.Rkk).ThenBy(x => x.Key).ToList();
                    SortMethod = "По количеству обращений";
                    break;
                case 3:
                    sorted_dict = dict.OrderByDescending(x => x.Value.Rkk).ThenBy(x => x.Value.Sum).ThenBy(x => x.Value.Obr).ThenBy(x => x.Key).ToList();
                    SortMethod = "По количеству РКК";
                    break;
                case 4:
                    sorted_dict = dict.OrderBy(x => x.Key).ThenBy(x => x.Value.Sum).ThenBy(x => x.Value.Obr).ThenBy(x => x.Value.Rkk).ToList();
                    SortMethod = "По фамилии ответственного исполнителя";
                    break;
            }
            return sorted_dict;
        }
    }
}
