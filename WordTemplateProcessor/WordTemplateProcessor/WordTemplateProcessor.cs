using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class WordTemplateProcessor
{
    public void InsertData(string templatePath, string outputPath, Dictionary<string, string> tagValues, Dictionary<string, Tuple<int, decimal, List<Tuple<string, string>>>> productsAndModules)
    {
        // Копируем шаблон в выходной файл
        File.Copy(templatePath, outputPath, true);

        // Открываем копию шаблона Word
        using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, true))
        {
            // Обрабатываем текстовые поля (Text Form Fields)
            foreach (var formField in doc.MainDocumentPart.Document.Descendants<SdtElement>())
            {
                foreach (var tag in tagValues)
                {
                    if (formField.SdtProperties.GetFirstChild<Tag>().Val == tag.Key)
                    {
                        formField.Descendants<Text>().First().Text = tag.Value;
                    }
                }
            }

            // Обрабатываем закладки (Bookmarks)
            foreach (var bookmarkStart in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                string bookmarkName = bookmarkStart.Name;
                if (tagValues.ContainsKey(bookmarkName))
                {
                    var parent = bookmarkStart.Parent;
                    if (parent != null)
                    {
                        var text = parent.Descendants<Text>().FirstOrDefault();
                        if (text != null)
                        {
                            text.Text = tagValues[bookmarkName];
                        }
                    }
                }
            }

            // Вставка данных продуктов и модулей
            InsertProductAndModuleData(doc.MainDocumentPart, "product_titles_and_m", "product_prices_and_m", "product_counts_and_m", productsAndModules);

            // Сохраняем изменения
            doc.Save();
        }
    }

    private void InsertProductAndModuleData(MainDocumentPart mainPart, string titlesBookmarkName, string pricesBookmarkName, string countsBookmarkName, Dictionary<string, Tuple<int, decimal, List<Tuple<string, string>>>> productsAndModules)
    {
        var titlesBookmark = mainPart.Document.Body.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == titlesBookmarkName);
        var pricesBookmark = mainPart.Document.Body.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == pricesBookmarkName);
        var countsBookmark = mainPart.Document.Body.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == countsBookmarkName);

        if (titlesBookmark != null && pricesBookmark != null && countsBookmark != null)
        {
            var titlesParent = titlesBookmark.Parent;
            var pricesParent = pricesBookmark.Parent;
            var countsParent = countsBookmark.Parent;

            foreach (var product in productsAndModules)
            {
                var productTitleRun = new Run(new Text(product.Key));
                var productPriceRun = new Run(new Text(product.Value.Item2.ToString("F2"))); // Цена продукта
                var productCountRun = new Run(new Text(product.Value.Item1.ToString())); // Количество копий продукта

                titlesParent.Append(productTitleRun.CloneNode(true));
                titlesParent.Append(new Break());
                pricesParent.Append(productPriceRun.CloneNode(true));
                pricesParent.Append(new Break());

                countsParent.Append(productCountRun.CloneNode(true));
                countsParent.Append(new Break());

                // Добавляем пустые строки между копиями
                for (int i = 0; i < product.Value.Item3.Count; i++)
                {
                    countsParent.Append(new Break());
                }

                foreach (var module in product.Value.Item3)
                {
                    var moduleTitleRun = new Run(new Text("\t" + module.Item1)); // Используем табуляцию для отступа
                    var modulePriceRun = new Run(new Text(module.Item2));

                    titlesParent.Append(moduleTitleRun.CloneNode(true));
                    titlesParent.Append(new Break());
                    pricesParent.Append(modulePriceRun.CloneNode(true));
                    pricesParent.Append(new Break());
                }
            }
        }
    }

    private void ReplaceBookmarkText(MainDocumentPart mainPart, string bookmarkName, string text)
    {
        var bookmarks = mainPart.Document.Body.Descendants<BookmarkStart>().Where(b => b.Name == bookmarkName);

        foreach (var bookmark in bookmarks)
        {
            var parent = bookmark.Parent;
            var runElement = bookmark.NextSibling<Run>();

            // Если Run элемент не существует, создаем его
            if (runElement == null)
            {
                runElement = new Run();
                parent.Append(runElement);
            }

            var textElement = runElement.GetFirstChild<Text>();

            // Если Text элемент не существует, создаем его
            if (textElement == null)
            {
                textElement = new Text();
                runElement.Append(textElement);
            }

            // Заменяем текст
            textElement.Text = text;
        }
    }
}
