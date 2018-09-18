/******************************************************************
** auth: wei.huazhong
** date: 9/17/2018 12:05:04 PM
** desc:
******************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService
{
    public class ImportError
    {
        public ImportError()
        { }
        public int Line
        {
            set;
            get;
        }

        public string ErrorMsg
        {
            set;
            get;
        }

        public string ErrorCode
        {
            set;
            get;
        }

        public List<int> ConflictLines
        {
            set;
            get;
        }

        public void FillConflictErrorMsg()
        {
            var linesBuilder = new StringBuilder();
            var lastLine = ConflictLines.Count - 1;
            for (int i = 0; i < lastLine; i++)
            {
                linesBuilder.Append(ConflictLines[i]);
                linesBuilder.Append(",");
            }
            linesBuilder.Append(ConflictLines[lastLine]);

            ErrorMsg = ErrorMessageHandler.Instance.GetErrorMessage(ErrorCode, linesBuilder.ToString());
        }

    }
}
