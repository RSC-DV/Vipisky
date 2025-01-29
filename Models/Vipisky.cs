using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Sebtum.Models
{
    public class DebtInfo
    {
        public string DebtOrderNumber1 { get; set; }
        public string DebtOrderNumber2 { get; set; }
        public DateTime DebtDate { get; set; }
    }
    public class Настройки
    {
        public int инд_Дата_проводки { get; set; }
        public int инд_Счет_Дебет { get; set; }
        public int инд_Сумма_Кредит { get; set; }
        public int инд_Номер_документа { get; set; }
        public int инд_Назначение_платежа { get; set; }
        public int инд_Строка_старта { get; set; }
        public Настройки()
        {
            Установить_настройки_по_умолчанию();
        }

        public Настройки(int Дата_проводки, int Счет_Дебет, int Сумма_Кредит, int Номер_документа, int Назначение_платежа)
        {
            инд_Дата_проводки = 1;
            инд_Счет_Дебет = 4;
            инд_Сумма_Кредит = 13;
            инд_Номер_документа = 14;
            инд_Назначение_платежа = 20;
        }

        public Настройки(DataTable Таблица)
        {
            Установить_настройки_автоматический(Таблица);
        }

        private void Установить_настройки_по_умолчанию()
        {
            инд_Дата_проводки = 1;
            инд_Счет_Дебет = 4;
            инд_Сумма_Кредит = 13;
            инд_Номер_документа = 14;
            инд_Назначение_платежа = 20;
        }

        private void Установить_настройки_автоматический(DataTable Таблица)
        {
        }
    }

    internal class Анализ_выписки
    {
        // Паттерн для поиска ФИО

        private const string Патерн_ФИО_Большими = @"[А-Я]+ [А-Я]+ [А-Я]+";

        private const string Патерн_ФИО_Универсальный = @"[А-Я]+[а-яё]*\s*([А-Я]+\s*[а-яё]*\s*[А-Я]+\s*[а-яё]*)*";
        private const string Патерн_ФИО_Корткое_Большими = @"[А-Я][а-яА-Я]+ [А-Я][а-яА-Я]+ [А-Я][а-яА-Я]+";
        private const string Патерн_ФИО = @"[А-Я][а-яё]+ [А-Я][а-яё]+ [А-Я][а-яё]+";
        private const string Патерн_ФИО_Короткое = @"[А-Я][а-яё]+ [А-Я]\. [А-Я]\.";

        // Паттерн для поиска периода
        private const string Патерн_Период = @"период 20\d\d\d\d";

        // Паттерн для поиска лицевого счета
        private const string Патерн_ЛС = @"13[0-9\-]+";

        // Паттерн для поиска адреса с кодом региона "68"

        List<string> patterns = new List<string>
{
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*ул\.\s*[^,]+,\s*д\.\s*\S+(?:,?\s*[А-Я]?)?,?\s*(?:кв\.\s*\S+|комн\.\s*\S+)?\b",
    @"68\d{4}[а-яА-Я0-9\.\,?\s?\-?]*\;?",
    @"68\d{4}[а-яА-Я0-9\.?,?\s?\-?]+\;? ",
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*ул\.\s*[^,]+,\s*д\.\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*ул\.\s*[^,]+,\s*д\.\s*\S+\b",
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*ул\.\s*[^,]+,\s*д\.\s*\S+,\s*кв\.\s*\S+,\s*комн\.\s*\S+\b",
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*ул\.\s*[^,]+,\s*д\.\s*\S+,\s*корп\.\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*ул\.\s*[^,]+,\s*д\.\s*\S+,\s*корп\.\s*\S+\b",
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+\b",  // Общее правило для улиц без указания "ул."
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*кв\.\s*\S+\b",  // Улица без указания "ул.", но с номером дома и квартиры
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*комн\.\s*\S+\b",  // Улица без указания "ул.", но с номером дома и комнаты
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*корп\.\s*\S+\b",  // Улица без указания "ул.", но с номером дома и корпуса
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*корп\.\s*\S+,\s*кв\.\s*\S+\b",  // Улица без указания "ул.", но с номером дома, корпуса и квартиры
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*корп\.\s*\S+,\s*комн\.\s*\S+\b",  // Улица без указания "ул.", но с номером дома, корпуса и комнаты
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*кв\.\s*\S+\b",  // Улица без указания "ул.", но с номером дома и квартиры
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*комн\.\s*\S+\b",  // Улица без указания "ул.", но с номером дома и комнаты
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*корп\.\s*\S+\b",  // Улица без указания "ул.", но с номером дома и корпуса
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*корп\.\s*\S+,\s*кв\.\s*\S+\b",  // Улица без указания "ул.", но с номером дома, корпуса и квартиры
    @"\b(\d{6}),\s*(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*корп\.\s*\S+,\s*комн\.\s*\S+\b",
    @"\b(\d{6})\s*,?\s*(г\.\s*[^,]+)\s*,?\s*(ул\.\s*[^,]+)\s*,?\s*(д\.\s*\S+)\s*,?\s*(кв\.\s*\S+)?\b", // Улица без указания "ул.", но с номером дома, корпуса и комнаты
    @"\b(г\.\s*[^,]+),\s*ул\.\s*[^,]+,\s*д\.\s*\S+(?:,?\s*[А-Я]?)?,?\s*(?:кв\.\s*\S+|комн\.\s*\S+)?\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*комн\.\s*\S+\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*корп\.\s*\S+\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*корп\.\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*корп\.\s*\S+,\s*комн\.\s*\S+\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*комн\.\s*\S+\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*корп\.\s*\S+\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*корп\.\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(г\.\s*[^,]+),\s*[^,]+,\s*д\.\s*\S+,\s*\S+,\s*корп\.\s*\S+,\s*комн\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*комн\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*корп\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*корп\.\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*корп\.\s*\S+,\s*комн\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*\S+,\s*комн\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*\S+,\s*корп\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*\S+,\s*корп\.\s*\S+,\s*кв\.\s*\S+\b",
    @"\b(ул\.\s*[^,]+),\s*д\.\s*\S+,\s*\S+,\s*корп\.\s*\S+,\s*комн\.\s*\S+\b"
};

        // private const string Патерн_Адрес = @"68\d{4}[а-яА-Я0-9\.?,?\s?\-?]+\;? ";
        private static string FindMatchingPattern(List<string> patterns, string text)
        {
            foreach (var pattern in patterns)
            {
                Match match = Regex.Match(text, pattern);
                if (match.Success)
                {
                    return match.Value.Trim();
                }
            }
            return " ";
        }
        public class DebtExtractor
        {
            private static List<(string pattern, int[] groupIndexes)> debtPatterns = new List<(string, int[])>
{
    // Patterns with 3 group indexes
    (@".*Судебный приказ (\d+-\d+/\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по и/п (\d+/\d+/\d+)-ИП", new int[] { 1, 3, 2 }),
    (@"По и/п (\d+/\d+/\d+)-ИП взыскан долг с .* Исполнительный лист ВС (\d+) от (\d{2}\.\d{2}\.\d{4})", new int[] { 1, 2, 3 }),
    (@"По и/п (\d+/\d+/\d+)-ИП взыскан долг с .* Судебный приказ (\d+-\d+/\d+/\d+) от (\d{2}\.\d{2}\.\d{4})", new int[] { 1, 2, 3 }),
    (@"По и/п (\d+/\d+/\d+)-ИП взыскан долг с .* от (\d{2}\.\d{2}\.\d{4}) Исполнительный лист ВС (\d+)", new int[] { 1, 3, 2 }),
    (@"Судебный приказ (\d+-\d+/\d+/\d+) от (\d{2}\.\d{2}\.\d{4}) по и/п (\d+/\d+/\d+)-ИП", new int[] { 1, 3, 2 }),
    (@"Исполнительный лист ВС (\d+) от (\d{2}\.\d{2}\.\d{4}) по и/п (\d+/\d+/\d+)-ИП", new int[] { 1, 3, 2 }),
    (@"Судебный приказ (\d+-\d+/\d+/\d+) от (\d{2}\.\d{2}\.\d{4}) по исполнительному листу (\d+/\d+/\d+)-ИП", new int[] { 1, 3, 2 }),
    (@"По и/п \d+/(\d+/\d+-\d+-\d+) Судебный приказ (\d+-\d+/\d+-\d+) от (\d{2}\.\d{2}\.\d{4})", new int[] { 1, 2, 3 }),
    (@"И/П (\d+/\d+/\d+)-ИП;.*Судебный приказ(?:№)?(\d+-\d+/\d+)(?:от)?(\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)", new int[] { 1, 2, 3 }),
    (@"(?:\([^)]+\))?(?:\s*\([^)]+\))?\s*.*Судебный приказ\s+(\d+-\d+/\d+)(?:\s*от\s*)?(\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)?\s*по\s*и/п\s*(\d+/\d+/\d+)-ИП", new int[] { 1, 2, 3 }),
    (@"Исполнительный лист ВС (\d+) по делу (\d+-\d+/\d+-\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)", new int[] { 1, 2, 3 }),
    (@"Судебный приказ (\d+-\d+/\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по делу (\d+-\d+/\d+-\d+)", new int[] { 1, 3, 2 }),
    (@"Исполнительный лист ВС (\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по делу (\d+-\d+/\d+-\d+)", new int[] { 1, 3, 2 }),
    (@"Исполнительный лист ВС (\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по делу (\d+-\d+/\d+-\d+)", new int[] { 1, 3, 2 }),
    (@"Исполнительный лист ВС (\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по делу (\d+-\d+/\d+-\d+)", new int[] { 1, 3, 2 }),
    (@"Судебный приказ №(\d+-\d+/\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по и/п (\d+/\d+/\d+)-ИП", new int[] { 1, 3, 2 }),
    (@"Судебный приказ (\d+-\d+/\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по и/п (\d+/\d+/\d+)-ИП", new int[] { 1, 3, 2 }),
    (@"Судебный приказ (\d+-\d+/\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по исполнительному листу (\d+/\d+/\d+)-ИП", new int[] { 1, 3, 2 }),
    (@"Исполнительный лист ВС (\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по и/п (\d+/\d+/\d+)-ИП", new int[] { 1, 3, 2 }),
    (@".*Судебный приказ (\d+-\d+/\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?) по и/п (\d+/\d+/\d+)-ИП", new int[] { 1, 2, 3 }),
    
    // Patterns with 2 group indexes
    (@"Долг по ИД (\d+-\d+/\d+/\d+) от (\d{2}\.\d{2}\.\d{4})", new int[] { 1, 2 }),
    (@"Взыскание по ИД от (\d{2}\.\d{2}\.\d{4}) №(\d+-\d+/\d+)", new int[] { 2, 1 }),
    (@"Судебный приказ (\d+-\d+/\d+-\d+) от (\d{2}\.\d{2}\.\d{4})", new int[] { 1, 2 }),
    (@"Исполнительный лист ВС (\d+) от (\d{2}\.\d{2}\.\d{4})", new int[] { 1, 2 }),
    (@"Судебный приказ (\d+-\d+/\d+/\d+) от (\d{2}\.\d{2}\.\d{4})", new int[] { 1, 2 }),
    (@"Взыскание по ИД от (\d{2}\.\d{2}\.\d{4}) ИД (\d+-\d+/\d+)", new int[] { 2, 1 }),
    (@"Исполнительный лист ВС (\d+) по и/п (\d+/\d+/\d+)-ИП", new int[] { 1, 2 }),
    (@"Долг по ИД\s+(\d+-\d+/\d+(-\d+)?)\s+от\s+(\d{2}\.\d{2}\.\d{4})", new int[] { 1, 3 }),
    (@"Исполнительный лист ВС (\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)", new int[] { 1, 2 }),
    (@"Исполнительный лист ВС (\d+) по и/п (\d+/\d+/\d+)-ИП", new int[] { 1, 2 }),
    (@"Судебный приказ (\d+-\d+/\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)", new int[] { 1, 2 }),
    (@"Судебный приказ (\d+-\d+/\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)", new int[] { 1, 2 }),
    (@"Исполнительный лист ВС (\d+) по делу (\d+-\d+/\d+-\d+)", new int[] { 1, 2 }),
    (@"Судебный приказ (\d+-\d+/\d+) от (\d{1,2}\.\d{1,2}(?:\.\d{2,4})?)", new int[] { 1, 2 }),
    
    // Patterns with 1 group index
    (@"По и/п (\d+/\d+/\d+)-ИП", new int[] { 1 }),
    (@"И/п (\d+/\d+/\d+)-ИП", new int[] { 1 }),
    (@"По делу (\d+-\d+/\d+-\d+)", new int[] { 1 }),
    (@"(\d+-\d+/\d+-\d+)", new int[] { 1 }),
};

            public static string ExtractDebtInfo(string text)
            {
                foreach (var pattern in debtPatterns)
                {
                    Match match = Regex.Match(text, pattern.pattern);
                    if (match.Success)
                    {
                        string orderNumber1 = match.Groups[pattern.groupIndexes[0]].Value;
                        string orderNumber2 = pattern.groupIndexes.Length > 2 ? match.Groups[pattern.groupIndexes[1]].Value : string.Empty;

                        // Handle partial date formats
                        string dateStr = match.Groups[pattern.groupIndexes[^1]].Value;
                        DateTime debtDate;
                        if (DateTime.TryParseExact(dateStr, new[] { "dd.MM.yyyy", "dd.MM.yy", "dd.MM" }, null, DateTimeStyles.None, out debtDate))
                        {
                            return $"{orderNumber1} | {orderNumber2} ({debtDate:dd.MM.yyyy})";
                        }
                        else
                        {
                            return $"{orderNumber1} | {orderNumber2} (Не удалось извлечь дату)";
                        }
                    }
                }

                return string.Empty; // If no match is found
            }
        }


        /* public static (bool hasDebt, DateTime? debtDate, string? debtOrderNumber) ExtractDebtInfo(string text)
         {
             bool hasDebt = false;
             DateTime? debtDate = null;
             string? debtOrderNumber = null;

             // Определение регулярных выражений для различных форматов записи долгов
             string debtOrderPattern1 = @"Долг по ИД (\d+-\d+/\d+/\d+) от (\d{2}\.\d{2}\.\d{4})";
             string debtOrderPattern2 = @"Взыскание по ИД от (\d{2}\.\d{2}\.\d{4}) №(\d+-\d+/\d+)";
             string debtOrderPattern3 = @"Судебный приказ (\d+-\d+/\d+-\d+) от (\d{2}\.\d{2}\.\d{4})";
             string debtOrderPattern4 = @"Исполнительный лист ФС (\d+) от (\d{2}\.\d{2}\.\d{4})";
             string debtOrderPattern5 = @"По и/п \d+/(\d+/\d+-\d+-\d+) Судебный приказ (\d+-\d+/\d+-\d+) от (\d{2}\.\d{2}\.\d{4})";
             string debtOrderPattern6 = @"Долг по ИД\s+(\d+-\d+/\d+(-\d+)?)\s+от\s+(\d{2}\.\d{2}\.\d{4})";

             // Сопоставление каждого шаблона с текстом
             Match match1 = Regex.Match(text, debtOrderPattern1);
             Match match2 = Regex.Match(text, debtOrderPattern2);
             Match match3 = Regex.Match(text, debtOrderPattern3);
             Match match4 = Regex.Match(text, debtOrderPattern4);
             Match match5 = Regex.Match(text, debtOrderPattern5);
             Match match6 = Regex.Match(text, debtOrderPattern6);
             // Проверка на успешное совпадение с любым из шаблонов
             if (match1.Success || match2.Success || match3.Success || match4.Success || match5.Success || match6.Success)
             {
                 hasDebt = true;

                 // Извлечение даты долга и номера приказа на основе успешного совпадения
                 if (match1.Success)
                 {
                     debtOrderNumber = match1.Groups[1].Value;
                     debtDate = DateTime.ParseExact(match1.Groups[2].Value, "dd.MM.yyyy", null);
                 }
                 else if (match2.Success)
                 {
                     debtDate = DateTime.ParseExact(match2.Groups[1].Value, "dd.MM.yyyy", null);
                     debtOrderNumber = match2.Groups[2].Value;
                 }
                 else if (match3.Success)
                 {
                     debtOrderNumber = match3.Groups[1].Value;
                     debtDate = DateTime.ParseExact(match3.Groups[2].Value, "dd.MM.yyyy", null);
                 }
                 else if (match4.Success)
                 {
                     debtOrderNumber = match4.Groups[1].Value;
                     debtDate = DateTime.ParseExact(match4.Groups[2].Value, "dd.MM.yyyy", null);
                 }
                 else if (match5.Success)
                 {
                     debtOrderNumber = match5.Groups[2].Value;
                     debtDate = DateTime.ParseExact(match5.Groups[3].Value, "dd.MM.yyyy", null);
                 }
                 else if (match6.Success)
                 {
                     debtOrderNumber = match6.Groups[1].Value;
                     debtDate = DateTime.ParseExact(match6.Groups[3].Value, "dd.MM.yyyy", null);

                 }
             }

             return (hasDebt, debtDate, debtOrderNumber);
         }*/

        private static string Найти(string Строка, string Патерн)
        {
            Match match = Regex.Match(Строка, Патерн, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Value;
            }
            return null;
        }

        public static string Найти_ФИО(string Строка)
        {

            Match match = Regex.Match(Строка, Патерн_ФИО_Корткое_Большими);
            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                Match match1 = Regex.Match(Строка, Патерн_ФИО_Большими);
                if (match1.Success)
                {
                    return match.Value;
                }
                else

                {
                    Match match2 = Regex.Match(Строка, Патерн_ФИО);
                    if (match2.Success)
                    {
                        return match2.Value;
                    }
                    else { return null; }
                }

            }



        }


        public DataTable Сделать_анализ(DataTable Выписка, Настройки настройки)
        {
            DataTable Новая_таблица = new DataTable();
            Новая_таблица.Columns.Add("ЛС", typeof(string));
            Новая_таблица.Columns.Add("ФИО", typeof(string));
            Новая_таблица.Columns.Add("Адрес", typeof(string));
            Новая_таблица.Columns.Add("Номер_документа", typeof(string));
            Новая_таблица.Columns.Add("Дата_проводки", typeof(string));
            Новая_таблица.Columns.Add("Сумма_кредит", typeof(double));

            Новая_таблица.Columns.Add("Период", typeof(string));
            Новая_таблица.Columns.Add("Приказ", typeof(string));

            // Новая колонка для объединения данных


            foreach (DataRow строка in Выписка.Rows)
            {
                try
                {
                    if (строка[настройки.инд_Сумма_Кредит].ToString() != "")
                    {
                        DataRow Новая_строка = Новая_таблица.NewRow();

                        // ЛС 
                        string ls = Найти(строка[настройки.инд_Назначение_платежа].ToString(), Патерн_ЛС);



                        // Дата проводки
                        double excelDate = Convert.ToDouble(строка[настройки.инд_Дата_проводки]);
                        DateTime dateTime = new DateTime(1899, 12, 30).AddDays(excelDate);
                        string датаПроводки = dateTime.ToString("dd.MM.yyyy");
                        //DateTime dateTime = new DateTime(1900, 1, 1).AddDays(Convert.ToDouble(строка[настройки.инд_Дата_проводки]));
                        Новая_строка["Дата_проводки"] = датаПроводки; // Дата без времени
                        Новая_строка["ЛС"] = ls;


                        // ФИО
                        var исключения = new List<string> { "ГОС ПОШЛИНА ВОЗМЕЩЕНИЕ", "ОПЛАТА ПО ДОГ", "УСЛУГИ ПО ОБРАЩЕНИЮ", "ОПЛАТА УСЛУГ ПО", "ОПЛАТА ЗА ВЫВОЗ", "ОПЛАТА ПО ДОГОВОРУ", "МИНИСТЕРСТВО ФИНАНСОВ ХАБАРОВСКОГО", "Лазо ГУФССП России", "Хабаровска ГУФССП России", "ОПЛАТА ПО СЧ", "Вывоз ТБО Без", "ОПЛАТА Без НДС", "ОПЛАТА ЗА УСЛУГИ" };

                        var фиоНазначение = Найти_ФИО(строка[настройки.инд_Назначение_платежа].ToString());
                        var фиоСчетДебет = Найти_ФИО(строка[настройки.инд_Счет_Дебет].ToString());

                        bool фиоНазначениеИсключение = СодержитИсключение(фиоНазначение, исключения);
                        bool фиоСчетДебетИсключение = СодержитИсключение(фиоСчетДебет, исключения);

                        if (фиоНазначение != null && !фиоНазначениеИсключение)
                        {
                            Новая_строка["ФИО"] = фиоНазначение;
                        }
                        else if (фиоСчетДебет != null && !фиоСчетДебетИсключение)
                        {
                            Новая_строка["ФИО"] = фиоСчетДебет;
                        }
                        else if (фиоНазначение != null)
                        {
                            Новая_строка["ФИО"] = фиоНазначение;
                        }
                        else if (фиоСчетДебет != null)
                        {
                            Новая_строка["ФИО"] = фиоСчетДебет;
                        }
                        else
                        {
                            Новая_строка["ФИО"] = " ";
                        }

                        // Функция для проверки наличия исключений в строке
                        bool СодержитИсключение(string текст, List<string> исключения)
                        {
                            if (текст == null) return false;
                            foreach (var исключение in исключения)
                            {
                                if (текст.Contains(исключение))
                                {
                                    return true;
                                }
                            }
                            return false;
                        }


                        // Адрес
                        Новая_строка["Адрес"] = FindMatchingPattern(patterns, (строка[настройки.инд_Назначение_платежа].ToString()));
                        // Сумма

                        Новая_строка["Сумма_кредит"] = Convert.ToDouble(строка[настройки.инд_Сумма_Кредит]);
                        double сумма = Convert.ToDouble(строка[настройки.инд_Сумма_Кредит]);
                        if (сумма > 10000)
                        { continue; }
                        else
                        {
                            var OOO = Найти(строка[настройки.инд_Счет_Дебет].ToString(), @"\bООО\b");
                            if (String.IsNullOrEmpty(ls) && OOO != null) { continue; }
                            else
                            {
                                Новая_строка["Номер_документа"] = строка[настройки.инд_Номер_документа].ToString();
                                // Период
                                Новая_строка["Период"] = Найти(строка[настройки.инд_Назначение_платежа].ToString(), Патерн_Период);

                                // Долг
                                Новая_строка["Приказ"] = DebtExtractor.ExtractDebtInfo(строка[настройки.инд_Назначение_платежа].ToString());


                                Console.WriteLine($"| {Новая_строка["Дата_проводки"],-15} | {Новая_строка["ФИО"],-10} | {Новая_строка["Адрес"],-15} | {Новая_строка["Сумма_кредит"],-15} | {Новая_строка["Номер_документа"],-15} | {Новая_строка["Период"],-10} | {Новая_строка["Приказ"],-15} |");
                                Новая_таблица.Rows.Add(Новая_строка);
                            }
                        }
                    }
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    continue;
                }
            }

            return Новая_таблица;
        }


    }
}