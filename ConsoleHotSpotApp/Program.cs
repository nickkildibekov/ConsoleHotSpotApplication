using System;

namespace ConsoleHotSpotApp
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            const int period = 1;
            var modifiedOnOrAfter = DateTime.Now.AddDays(-period);
            Methods.A(modifiedOnOrAfter);
        }
    }
}
