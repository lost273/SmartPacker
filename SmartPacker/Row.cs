using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartPacker {
    [Serializable]
    class Row {
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public int Quantity { get; set; }
        public decimal Cost { get; set; }
        public decimal Total { get; set; }
    }
}
