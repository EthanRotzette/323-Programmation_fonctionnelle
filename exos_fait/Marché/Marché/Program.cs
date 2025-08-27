using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Google.Protobuf.WellKnownTypes;
using IronXL;

namespace Marché
{
    internal class Program
    {
        static void Main(string[] args)
        {
            const string filePath = "C:\\Users\\pk50gbi\\Documents\\GitHub\\323-Programmation_fonctionnelle\\exos_fait\\Marché\\Place du marché.xlsx";

            if (File.Exists(filePath))
            {
                WorkBook workBook = WorkBook.Load(filePath); // charge le fichier excel
                WorkSheet workSheet = workBook.WorkSheets[1];
                List<string> listOfWatermelon = new List<string>();
                Dictionary<string, int> map = new Dictionary<string, int>();
                int sellerNumber = 0;

                foreach (var cell in workSheet["C2:C75"])
                {
                    //Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);

                    if (cell.Text == "Pastèques")
                    {
                        listOfWatermelon.Add(cell.AddressString);
                    }

                    if (cell.Text == "Pêches")
                    {
                        sellerNumber++;
                    }
                }

                Console.WriteLine("Il y a " + sellerNumber + " vendeurs de pêches");

                foreach (string address in listOfWatermelon)
                {
                    //besoins des infos de la colonnes A et D
                    string num = address.Substring(1); // retire la deuxième lettre

                    //string Seller = workSheet["B" + num].StringValue;
                    int NumberWatermelon = workSheet["D" + num].IntValue;

                    map.Add(num, NumberWatermelon);
                    //Console.WriteLine(NumberWatermelon);
                    //Console.WriteLine(Seller + " vend " + NumberWatermelon + " pastèques au stand " + Emplacement);
                }

                int maxWater = -1;
                string bestProducer = "oopssss";
                int Emplacement = 0;
                int NumberWaterMelon = 0;
                foreach(KeyValuePair<string,int> entry in map)
                {
                    if(entry.Value > maxWater)
                    {
                        maxWater = entry.Value;
                        bestProducer = workSheet["B" + entry.Key].StringValue;
                        Emplacement = workSheet["A" + entry.Key].IntValue;
                        NumberWaterMelon = workSheet["D" + entry.Key].IntValue;
                    }
                }

                Console.WriteLine("C'est " + bestProducer + " qui a le plus de pastèque (stand " + Emplacement + ", " + NumberWaterMelon + " pièces)");

                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Le fichier n'est pas la !");
                Console.ReadLine();
            }

        }
    }
}
