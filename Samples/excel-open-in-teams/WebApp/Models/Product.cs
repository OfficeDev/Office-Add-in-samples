using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApp.Models
{
    public class Product
    {
       public int ID { get; set; }
        public string Name { get; set; }
        public int Qtr1 { get; set; }
        public int Qtr2 { get; set; }
        public int Qtr3 { get; set; }
        public int Qtr4 { get; set; }

    }
}