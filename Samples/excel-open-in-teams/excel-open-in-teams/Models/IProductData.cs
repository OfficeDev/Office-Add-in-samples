using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace excel_open_in_teams.Models
{
    public interface IProductData
    {
        IEnumerable<Product> GetAll();
    }

    /// <summary>
    /// For dev testing creates an in memory database with some test product data
    /// </summary>
    public class InMemoryProductData : IProductData
    {
        List<Product> products;
        public InMemoryProductData()
        {
            products = new List<Product>()
            { new Product {ID=1, Name="Frames", Qtr1=5000, Qtr2=7000, Qtr3=6544, Qtr4=4377},
            new Product {ID=2, Name="Saddles", Qtr1=400, Qtr2=323, Qtr3=276, Qtr4=651},
            new Product {ID=3, Name="Brake levers", Qtr1=12000, Qtr2=8766, Qtr3=8456, Qtr4=9812},
            new Product {ID=4, Name="Chains", Qtr1=1550, Qtr2=1088, Qtr3=692, Qtr4=853},
            new Product {ID=5, Name="Mirrors", Qtr1=225, Qtr2=600, Qtr3=923, Qtr4=544},
            new Product {ID=5, Name="Spokes", Qtr1=6005, Qtr2=7634, Qtr3=4589, Qtr4=8765}
            };
        }

        public IEnumerable<Product> GetAll()
        {
            return from r in products
                   orderby r.ID
                   select r;
        }
    }
}
