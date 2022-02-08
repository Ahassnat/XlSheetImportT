using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace DotnetXlSheetImportTamer.Models
{
    public class NVPCisco
    {
        [Key]
        [Required]
        public string PartSKU { get; set; }
        [Required]
        public string CategoryCode { get; set; }
        [Required]
        public string Brand { get; set; }
        [Required]
        public string Manufacturer { get; set; }
        [Required]
        public string ItemDescription { get; set; }
        [Required]
        public string PriceList { get; set; }
        [Required]
        public string MinDiscount { get; set; }
        [Required]
        public string DiscountPrice { get; set; }
    }
}
