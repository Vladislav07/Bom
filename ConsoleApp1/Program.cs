using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
  using AccessToBom_Cube;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] result = null;
            string pathFile = @"A:\My\10-тонный электрический подъемный кран\ramka-1.SLDDRW";
            ISpec sp = new Spec();
            result = sp.GetListBom(pathFile);

            Console.Write(result[0]);
            Console.Read();
        }
    }
}
