using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace raskroiParser1.components
{
    public static class tableParsing
    {
        public static string tableParsingFunc(Microsoft.Office.Interop.Word.Table table)
        {
            bool headerOftable = true;
            List<string> Header = new List<string>();
            string jsonDoc = @"[{";

            foreach (Microsoft.Office.Interop.Word.Row row in table.Rows)
            {
                int indexOfHeader = 0;
                int countOfCells = row.Cells.Count;
                foreach (Microsoft.Office.Interop.Word.Cell cell in row.Cells)
                {
                    string curCellVal = cell.Range.Text.ToString();

                    delEscapeChars.delEscapeCharsFunc(ref curCellVal);

                    if (String.IsNullOrEmpty(curCellVal))
                    {
                        curCellVal = "NotDefined";
                    }

                    if (headerOftable)
                    {
                        Header.Add(curCellVal);
                    }
                    else
                    {
                        jsonDoc = jsonDoc + "\"" + Header[indexOfHeader] + "\":\"" + curCellVal + "\"";
                        if (countOfCells != 1)
                        {
                            jsonDoc = jsonDoc + ",";
                        }
                        else
                        {
                            jsonDoc = jsonDoc + "}";
                        }
                    }
                    indexOfHeader++;
                    countOfCells--;
                }

                if (row.Next != null && headerOftable == false)
                {
                    jsonDoc = jsonDoc + ",{";
                }
                else if (row.Next == null && headerOftable == false)
                {
                    jsonDoc = jsonDoc + "]";
                }

                headerOftable = false;
            }

            return jsonDoc;
        }
    }
}
