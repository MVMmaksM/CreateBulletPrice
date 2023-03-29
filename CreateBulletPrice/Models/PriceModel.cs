using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateBulletPrice.Models
{
    internal abstract class PriceModel
    {
        public string? Name { get; set; }      
        public decimal? Ufa { get; set; }
        public decimal? Ijevsk { get; set; }
        public decimal? Perm { get; set; }
        public decimal? Orenburg { get; set; }
        public decimal? Kurgan { get; set; }
        public decimal? Ekaterinburg { get; set; }
        public decimal? Tumen { get; set; }
        public decimal? Hanty { get; set; }
        public decimal? Salehard { get; set; }
        public decimal? Chelyabinsk { get; set; }
    }
}
