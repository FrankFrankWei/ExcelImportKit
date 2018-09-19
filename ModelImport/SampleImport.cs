/******************************************************************
** auth: wei.huazhong
** date: 9/18/2018 12:00:21 PM
** desc:
******************************************************************/

using ExcelService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelImport
{
    public class SampleImport : ImportEntityBase
    {
        public SampleImport()
        { }

        public int Age { get; set; }
        public string Name { get; set; }
        public DateTime Birthday { get; set; }
        public float Height { get; set; }
        public decimal Money { get; set; }
        public bool Gender { get; set; }
        public string GenderName => Gender ? "Male" : "Female";
        /// <summary>
        /// 1: Student  2: Staff 3: Soldier
        /// </summary>
        public int State { get; set; }
        public string StateName { get { return State == 1 ? "Student" : State == 2 ? "Staff" : State == 3 ? "Soldier" : "not set"; } }
    }
}
