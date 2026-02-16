using System;

namespace Group4333.Models
{
    public class Service
    {
        public int Id { get; set; }           // ID услуги
        public string Name { get; set; }      // Название услуги
        public string Type { get; set; }      // Вид услуги (критерий группировки)
        public decimal Price { get; set; }     // Стоимость

        // Для удобного отображения в списке
        public override string ToString()
        {
            return $"{Name} - {Price} руб. ({Type})";
        }
    }
}