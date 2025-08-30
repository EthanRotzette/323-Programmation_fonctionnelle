using System;
using System.IO;
using IronXL;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;
using System.Collections.Generic;


namespace Marché
{
    internal class Program
    {
        const string filePath = "C:\\Users\\etrot\\OneDrive\\Documents\\GitHub\\323-Programmation_fonctionnelle\\exos_fait\\Marché\\Place du marché.xlsx"; //"C:\\Users\\pk50gbi\\Documents\\GitHub\\323-Programmation_fonctionnelle\\exos_fait\\Marché\\Place du marché.xlsx";
        static void Main(string[] args)
        {


            //******* V1 ********\\

            /*if (File.Exists(filePath))
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
            }*/

            ///********** V3 ********** \\\

            List<Seller> seller = new List<Seller>();

            if (File.Exists(filePath))
            {
                seller = Oui(seller);

                int nbrSellerPeach = seller.Where(s => s.ProductName == "Pêches").Count();

                seller = seller.Where(s => s.ProductName == "Pastèques").ToList();
                var bestSeller = seller.OrderByDescending(s => int.Parse(s.NbrProduct)).First();

                Console.WriteLine($"Il y a {nbrSellerPeach} vendeurs de pêches");
                Console.WriteLine($"C'est {bestSeller.SellerName} qui a le plus de pastèque (stand {bestSeller.Stand}, {bestSeller.NbrProduct} pièces)");
            }
            else
            {
                Console.WriteLine("Le fichier n'est pas la !");
                Console.ReadLine();
            }
        }

        class Seller
        {
            public string Stand { get; set; }
            public string SellerName { get; set; }
            public string ProductName { get; set; }
            public string NbrProduct { get; set; }
            public string QuantityName { get; set; }
            public string price { get; set; }
        }

        private static List<Seller> Oui(List<Seller> seller)
        {
            WorkBook workBook = WorkBook.Load(filePath); // charge le fichier excel
            WorkSheet workSheet = workBook.WorkSheets[1];

            // On parcourt chaque ligne du tableau (de 2 à 75)
            for (int row = 2; row <= 75; row++)
            {
                // On lit les colonnes de A à F
                string Stand = workSheet["A" + row].StringValue;
                string SellerName = workSheet["B" + row].StringValue;
                string ProductName = workSheet["C" + row].StringValue;
                string NbrProduct = workSheet["D" + row].StringValue;
                string QuantityName = workSheet["E" + row].StringValue;
                string Price = workSheet["F" + row].StringValue;

                // Si la ligne est vide, on saute
                if (string.IsNullOrWhiteSpace(SellerName) && string.IsNullOrWhiteSpace(ProductName))
                    continue;

                // On ajoute à la liste
                seller.Add(new Seller
                {
                    Stand = Stand,
                    SellerName = SellerName,
                    ProductName = ProductName,
                    NbrProduct = NbrProduct,
                    QuantityName = QuantityName,
                    price = Price
                });

            }
            return seller;
        }
    }
}
