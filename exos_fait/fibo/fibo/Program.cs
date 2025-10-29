for (int i = 0; i < 10; i++)
{
    Console.WriteLine($"{i} : {Fibonacci(i)}");
}

for (int i = 0; i < 93; i++)
{
    Console.WriteLine($"{i} : {FibonacciOptimise(i)}");
}

static int Fibonacci(int n)
{
    if (n >= 2)
        return Fibonacci(n - 1) + Fibonacci(n -2);

    return n;
}


static long FibonacciOptimise(long sequence, long precedent = 0, long actuel = 1)
{
    if (sequence > 0 )
        return FibonacciOptimise(sequence - 1,actuel,actuel+precedent);

    return precedent;
}