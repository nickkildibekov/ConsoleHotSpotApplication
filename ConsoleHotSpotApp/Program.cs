using System;

namespace ConsoleHotSpotApp
{
    internal class Program
    {
        private static void Main()
        {
            Console.Write("Enter a correct ( >= 0 )NUMBER  of days for analyst period :  ");
            var period = Convert.ToInt32(Console.ReadLine());
            if (period <= 0) Main();
            var modifiedOnOrAfter = DateTime.Now.AddDays(-period);
            Methods.A(modifiedOnOrAfter);
        }
    }
}

