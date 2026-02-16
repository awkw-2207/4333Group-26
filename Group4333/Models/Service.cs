using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Group4333.Models
{
    public class Service
    {
        public int Id { get; set; }        
        public string Name { get; set; }     
        public string Type { get; set; }   
        public decimal Price { get; set; }     

        public override string ToString()
        {
            return $"{Name} - {Price} руб. ({Type})";
        }
    }
}
