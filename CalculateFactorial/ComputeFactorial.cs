using System;
using System.Diagnostics;
using System.Numerics;

namespace CalculateFactorial
{
	class ComputeFactorial
	{
		static void Main(string[] args)
		{
			Console.Write("Enter number. 1 <= number <= 1000 \nnumber = ");
			int num = int.Parse(Console.ReadLine());
			if (num < 0 || num > 1000)
			{
				Console.WriteLine("Wrong input.");
				return;
			}
			Console.WriteLine();
			Console.WriteLine("Calculate factorial");
			Console.WriteLine("The result of {0}! is: {1}\n", num, CalculateFactorial(num));
			Console.WriteLine("Calculate factorial optimized");
			Console.WriteLine("The result of {0}! is: {1}\n", num, CalculateFactorialOptimized(num));
		}
		static BigInteger CalculateFactorial(int number)
		{
			BigInteger result = 1;
			if (number == 0)
			{
				return 1;
			}
			Stopwatch stopwatch = new Stopwatch();
			stopwatch.Start();
			for (int i = 1; i <= number; i++)
			{
				result *= i;
			}
			stopwatch.Stop();
			Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);
			return result;
		}

		//optimized factorial calculation
		static BigInteger CalculateFactorialOptimized(int number)
		{
			BigInteger sum = number;
			BigInteger result = number;
			if (number == 0)
			{
				return 1;
			}
			Stopwatch stopwatch = new Stopwatch();
			stopwatch.Start();
			for (int i = number - 2; i > 1; i -= 2)
			{
				sum = (sum + i);
				result *= sum;
			}

			if (number % 2 != 0)
			{
				result *= number / 2 + 1;
			}
			stopwatch.Stop();
			Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);
			return result;
		}
	}

}
