using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Words
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string[] words = { "bonjour", "hello", "monde", "vert", "rouge", "bleu", "jaune" };

            Func<string, bool> NoXandMoreThanFour = (word) => !word.Contains("x") && word.Length > 4 && word.Length == Math.Round(words.Average(w => w.Length), 0);
            //Func<string, bool> MoreThanFour = word => word.Length > 4;
            Action<string> writeToConsole = word => { Console.WriteLine(word); };
            
            words = words.Where(NoXandMoreThanFour).ToArray();
            //string[] wordsWhithoutXandMoreThanFour = wordsWhithoutX.Where(MoreThanFour).ToArray();

            words.OrderByDescending(w => w).ToList().ForEach(message => writeToConsole(message));
            Array.Sort(words);
            words.ToList().ForEach(message => writeToConsole(message));
            Array.Reverse(words);
            words.ToList().ForEach(message => writeToConsole(message));
            Console.ReadLine();
            
        }
    }
}
