using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace COFCO.SharedEntities.Models
{
    public class CofcoRowModel
    {
        public string Id { get; set; }
        public string Port { get; set; }
        public string Supplier { get; set; }
        public string Product { get; set; }
        public string Quantity { get; set; }
        public string Date { get; set; }
        public string VehicleNumber { get; set; }
        public string TTNNumber { get; set; }
        public string Contract { get; set; }
    }
}
