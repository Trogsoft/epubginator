using System;
using Trogsoft.CommandLine;

namespace epubginator
{
    class Program
    {
        static int Main(string[] args)
        {

            var parser = new Parser();
            return parser.Run(args);

        }
    }
}
