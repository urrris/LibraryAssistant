using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibraryAssistant.Models
{
    internal class Fact
    {
        public string FType { get; set; }
        public string TakingDate { get; set; }
        public string ReturnDate { get; set; }
        public string FactDate {  get; set; }
        public string Book {  get; set; }
        public Fact(string type, string takingDate, string returnDate, string factDate, string book) { 
            FType = type;
            TakingDate = takingDate;
            ReturnDate = returnDate;
            FactDate = factDate;
            Book = book;
        }
    }
}
