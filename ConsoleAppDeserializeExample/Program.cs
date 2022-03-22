using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppDeserializeExample
{
  internal class Program
  {
    static void Main()
    {
      Action<string> Display = Console.WriteLine;
      string excelFileName = "book1.xlsx";
      Display($"Serialize the {excelFileName} workbook");


      Display("Press any key to exit:");
      Console.ReadKey();
    }
  }
}
