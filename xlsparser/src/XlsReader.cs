using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace xlsparser
{
    class XlsReader : Singleton<XlsReader>
    {
        public bool ReadExcel(string path, List<ISheet> sheet_list)
        {
            bool is_succ = false;

            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook = null;
                    if (Path.GetExtension(fs.Name) == ".xls")
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    else if (Path.GetExtension(fs.Name) == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(fs);
                    }

                    if (null == workbook || workbook.NumberOfSheets < 1)
                    {
                        is_succ = false;
                    }
                    else
                    {
                        is_succ = true;
                        BaseXlsParser.formulaEvaluator = new HSSFFormulaEvaluator(workbook);

                        for (int i = 0; i < workbook.NumberOfSheets; i++)
                        {
                            ISheet sheet = workbook.GetSheetAt(i);
                            sheet_list.Add(sheet);
                            is_succ = true;
                        }
                    }
                }
            }
            catch (Exception)
            {
                is_succ = false;
            }

            return is_succ;
        }
    }
}
