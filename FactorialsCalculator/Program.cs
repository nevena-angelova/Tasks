using System;
using System.Diagnostics;

namespace FactorialsCalculator
{
    class Program
    {
        static void Main()
        {
            var n = 10;

            //Console.WriteLine(CalculateFactoriel(n));

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            var repeat = 1000;
            for (int i = 0; i < repeat; i++)
            {
                CalculateFactoriel(n);
            }

            stopwatch.Stop();
            Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);
        }

        static int CalculateFactoriel(int n)
        {
            var result = 1;

            if (n > 1)
            {
                for (int i = 2; i <= n; i++)
                {
                    result *= i;
                }
            }

            return result;
        }

    }
}
