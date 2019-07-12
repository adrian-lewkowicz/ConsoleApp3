using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {

            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(@"C:/Users/LENOVO/Desktop/excele/101102.xlsx", FileMode.Open, FileAccess.Read))
            {

                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheetAt(0);
           // Console.WriteLine(sheet.GetColumnWidth(row));
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    //Console.WriteLine(sheet.GetColumnWidth(row));
                    for (int c=0; c<sheet.GetColumnWidth(row); c++)
                    {
                        try
                        {
                            if (sheet.GetRow(row).GetCell(c) != null)
                            {
                                Console.Write(sheet.GetRow(row).GetCell(c).StringCellValue);
                            }
                        }
                        catch
                        {

                        }
                    }
                }
            }
            Console.ReadKey();
        }
    }
}
