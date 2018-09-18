/******************************************************************
** auth: wei.huazhong
** date: 9/18/2018 12:00:21 PM
** desc:
******************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelImport
{
    public class SampleImport : ImportBase
    {
        public SampleImport()
        { }

        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime Birthday { get; set; }
        public float Height { get; set; }
        public double Money { get; set; }
        public bool IsMale { get; set; }
    }
}
