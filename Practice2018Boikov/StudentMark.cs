using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practice2018Boikov
{
    class StudentMark
    {
        public int id { get; set; }
        public string last_name_ukr { get; set; }
        public string name_ukr { get; set; }
        public string group_number { get; set; }
        public string short_name { get; set; }
        public string name { get; set; }
        public string check_form { get; set; }
        public string name_1 { get; set; }
        public string last_name_ukr_1 { get; set; }
        public string name_ukr_1 { get; set; }
        public string chair_number { get; set; }
        public string chair_number_1 { get; set; }

        public StudentMark(int Id, string lastNameukr, string nameukr, string group, string shortName,
            string namec, string checkForm, string name1, string lastName1_ukr, string name1_ukr,
            string chairNumber, string chairNumber1)
        {
            id = Id;
            last_name_ukr = lastNameukr;
            name_ukr = nameukr;
            group_number = group;
            short_name = shortName;
            name = namec;
            check_form = checkForm;
            name_1 = name1;
            last_name_ukr_1 = lastName1_ukr;
            last_name_ukr_1 = name1_ukr;
            chair_number = chairNumber;
            chair_number_1 = chairNumber1;
        }

    }
}
