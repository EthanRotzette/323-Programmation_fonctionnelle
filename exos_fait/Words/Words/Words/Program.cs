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
            Action<string> writeToConsole = word => { Console.WriteLine(word); };
            /// Exercice A ///
            /*
            string[] words = { "bonjour", "hello", "monde", "vert", "rouge", "bleu", "jaune" };

            Func<string, bool> NoXandMoreThanFour = (word) => !word.Contains("x") && word.Length > 4 && word.Length == Math.Round(words.Average(w => w.Length), 0);
            //Func<string, bool> MoreThanFour = word => word.Length > 4;
            
            words = words.Where(NoXandMoreThanFour).ToArray();
            //string[] wordsWhithoutXandMoreThanFour = wordsWhithoutX.Where(MoreThanFour).ToArray();

            words.OrderByDescending(w => w).ToList().ForEach(message => writeToConsole(message));
            Array.Sort(words);
            words.ToList().ForEach(message => writeToConsole(message));
            Array.Reverse(words);
            words.ToList().ForEach(message => writeToConsole(message));
            Console.ReadLine();
            */
            /// Exercice B ///
            string[] words = { "whatThe!!!", "bonjour", "hello", "monde", "vert", "rouge", "bleu", "jaune", "My kingdom for a horse !", "Ooops I did it again" };

            words = words.Skip(1).Reverse().Skip(2).Reverse().ToArray();
            words.ToList().ForEach(message => writeToConsole(message));
            Console.ReadLine();
        }
    }
}
