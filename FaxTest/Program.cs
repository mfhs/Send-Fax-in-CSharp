using System;

namespace FaxTest
{
    class Program
    {
        static void Main(string[] args)
        {
            FaxSender fs = new FaxSender();
            fs.SendFax(); 
            Console.ReadLine();
        }
    }
}
