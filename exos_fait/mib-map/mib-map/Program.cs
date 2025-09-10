using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace mib_map
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var i18n = new Dictionary<string, string>()
            {
                { "Pommes","Apples"},
                { "Poires","Pears"},
                { "Pastèques","Watermelons"},
                { "Melons","Melons"},
                { "Noix","Nuts"},
                { "Raisin","Grapes"},
                { "Pruneaux","Plums"},
                { "Myrtilles","Blueberries"},
                { "Groseilles","Berries"},
                { "Tomates","Tomatoes"},
                { "Courges","Pumpkins"},
                { "Pêches","Peaches"},
                { "Haricots","Beans"}
            };

            List<Product> products = new List<Product>
            {
                new Product { Location = 1, Producer = "Bornand", ProductName = "Pommes", Quantity = 20,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 1, Producer = "Bornand", ProductName = "Poires", Quantity = 16,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 1, Producer = "Bornand", ProductName = "Pastèques", Quantity = 14,Unit = "pièce", PricePerUnit = 5.50 },
                new Product { Location = 1, Producer = "Bornand", ProductName = "Melons", Quantity = 5,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 2, Producer = "Dumont", ProductName = "Noix", Quantity = 20,Unit = "sac", PricePerUnit = 5.50 },
                new Product { Location = 2, Producer = "Dumont", ProductName = "Raisin", Quantity = 6,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 2, Producer = "Dumont", ProductName = "Pruneaux", Quantity = 13,Unit = "kg", PricePerUnit = 5.50 },
                new Product { Location = 2, Producer = "Dumont", ProductName = "Myrtilles", Quantity = 12,Unit = "kg", PricePerUnit = 5.50 },
            };

            // var result = products.Select(p => $"{p.Producer.Substring(0, 3)}...{p.Producer.Last()} {p.ProductName = i18n[p.ProductName]} {p.Quantity * p.PricePerUnit}");
            // Console.WriteLine("Seller | Product | CA");
            // result.ToList().ForEach(p => Console.WriteLine(p));
            // Console.ReadLine();

            //var resultCSV = products.Select(p => $"{p.Producer.Substring(0, 3)}...{p.Producer.Last()},{p.ProductName = i18n[p.ProductName]},{p.Quantity * p.PricePerUnit}");
            //StreamWriter streamWriter = new StreamWriter("file.csv"); // créer dans \bin\Debug
            //streamWriter.WriteLine("Seller,Product,CA");
            //resultCSV.ToList().ForEach(r => streamWriter.WriteLine(r));
            //streamWriter.Close();

            ///// Dashbord /////
            
            //var result = products.Select(p => (Nom: p.Producer.Substring(0, 1) + (p.Producer.Length - 1) + p.Producer.Last(), NomProduit: p.ProductName, 
            //Stock: p.Quantity < 10 ? "Stock faible" : p.Quantity >= 10 && p.Quantity <= 15 ? "Stock normal" : "stock élevé", 
            //Prix: p.Quantity < 10 ? (15 * p.PricePerUnit / 100) + p.PricePerUnit : p.Quantity >= 10 && p.Quantity <= 15 ? (5 * p.PricePerUnit / 100) + p.PricePerUnit : p.PricePerUnit, 
            //rentable : p.PricePerUnit * p.Quantity > 100 ? "Premium" : "Standard"));

            var tuple = (
                Nom: products.Select(p => p.Producer.Substring(0, 1) + (p.Producer.Length - 1) + p.Producer.Last()), 
                Stock: products.Select(p => p.Quantity < 10 ? "Stock faible" : p.Quantity >= 10 && p.Quantity <= 15 ? "Stock normal" : "stock eleve"),
                Prix: products.Select(p => p.Quantity < 10 ? Math.Round((15 * p.PricePerUnit / 100) + p.PricePerUnit) : p.Quantity >= 10 && p.Quantity <= 15 ? Math.Round((5 * p.PricePerUnit / 100) + p.PricePerUnit) : p.PricePerUnit),
                rentable: products.Select(p => p.PricePerUnit * p.Quantity > 100 ? "Premium" : "Standard")
                );

            var options = new JsonSerializerOptions { IncludeFields = true };
            string json = JsonSerializer.Serialize(tuple, options);
            File.WriteAllText("file.json",json);
        }

        class Product
        {
            public int Location { get; set; }
            public string Producer { get; set; }
            public string ProductName { get; set; }
            public int Quantity { get; set; }
            public string Unit { get; set; }
            public double PricePerUnit { get; set; }
        }
    }
}
