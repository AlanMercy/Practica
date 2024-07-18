using System;
using System.Collections.Generic;

public class Program
{
    public static void Main()
    {
        // Путь к исходному шаблону Word
        string templatePath = "Template.docx";

        // Путь для сохранения измененного шаблона Word
        string outputPath = "Output.docx";

        // Строка подключения к базе данных
        string connectionString = "your_connection_string_here";

        try
        {
            // Пример данных для замены тегов
            var tagValues = new Dictionary<string, string>
            {
                { "inf_comp", "Компания ООО" },
                { "inf_comp_2", "Информация о компании" },
                { "inf_client", "Информация о клиенте" },
                { "comp_inf", "Дополнительная информация о компании" },
                { "kom_id", "1" },
                { "kom_date", "2024-07-16" },
                { "order_price1", "460.00" },
                { "order_price2", "460.00" },
                { "order_discount", "20.00" },
                { "order_difference", "440.00" }
            };

            // Пример данных для продуктов и модулей
            var productsAndModules = new Dictionary<string, Tuple<int, decimal, List<Tuple<string, string>>>>
            {
                {
                    "Продукт 1", Tuple.Create(2, 100.00m, new List<Tuple<string, string>>
                    {
                        Tuple.Create("Модуль 1.1", "50.00"),
                        Tuple.Create("Модуль 1.2", "30.00")
                    })
                },
                {
                    "Продукт 2", Tuple.Create(1, 90.00m, new List<Tuple<string, string>>
                    {
                        Tuple.Create("Модуль 2.1", "70.00"),
                        Tuple.Create("Модуль 2.2", "20.00")
                    })
                }
            };

            // Создаем экземпляр WordTemplateProcessor
            var templateProcessor = new WordTemplateProcessor();

            // Вызываем метод для вставки данных в шаблон
            templateProcessor.InsertData(templatePath, outputPath, tagValues, productsAndModules);

            Console.WriteLine("Данные успешно вставлены в шаблон Word.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Произошла ошибка при вставке данных в шаблон: {ex.Message}");
        }
    }
}