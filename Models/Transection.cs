using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace TestNagis.Models
{
    [Table("Transections")]
    public class Transection
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public string Buyer { get; set; }
        public string Seller { get; set; }
        public double Amount { get; set; }
        public DateTime Date { get; set; }
    }
}