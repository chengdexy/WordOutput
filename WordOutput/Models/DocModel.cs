using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordOutput.Models {
    class DocModel {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Gender { get; set; }
        public string Age { get; set; }
        public string SickYears { get; set; }
        public string PhoneNumbers { get; set; }
        public string Address { get; set; }
        public string VipNumber { get; set; }
        public string KindOfSick { get; set; }
        public string UsingDrugs { get; set; }
        public string CurSymptoms { get; set; }//目前症状
        public string Records { get; set; }
    }
}
