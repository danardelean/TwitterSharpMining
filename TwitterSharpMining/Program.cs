using System;

namespace TwitterSharpMining
{
    class Program
    {
        static void Main(string[] args)
        {
            TwitterHelper.DumpFollowers("evoespueblo", new DateTime(2019, 10, 20), -1, 0, 0);
        }
    }
}
