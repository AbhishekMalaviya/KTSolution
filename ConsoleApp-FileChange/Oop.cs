using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp_FileChange
{
    public record recordTest(string name)
    {

    }

    public class clsB
    {
        private int _number;
        private string _name;
        public clsB(string name)
        {
            Name = name;
        }



        public void SetValue(int number, ref string name)
        {
            _number = number + 1;
            _name = name = name + " Change";
        }

        public void DisplayValue()
        {
            Console.WriteLine($"{_number} {_name}");
        }


        public string Name { get; set; }
    }
    internal class BaseClass
    {
        internal virtual void DisplayMessage(string msg)
        {
            Console.WriteLine($"Base class- {msg}");
        }
    }

    internal class ChildClass : BaseClass
    {
        internal override void DisplayMessage(string msg)
        {
            Console.WriteLine($"Child class- {msg}");
        }
    }
}
