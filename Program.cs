using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

internal class Program
{
    // ============================================================
    // BEÁLLÍTÁSOK
    // ============================================================

    private const int TotalOwners = 3_640_450;

    // A korábban meghatározott becslés alapján:
    private const int NaturalPersonCount = 3_391_650;
    private const int LegalPersonCount = 248_800;

    // Az eredeti járműállomány:
    private const int NaturalPersonVehicles = 4_470_202;
    private const int LegalPersonVehicles = 870_634;

    private const double SecondFirstNameProbability = 0.10;
    private const double SecondLastNameProbability = 0.10;

    private const string NamesDirectory =
        @"D:\Egyetem\7+félév\Szakdolgozat II\generalo algoritmusok\autótulaj generálás\nevek";

    private static readonly string LastNamesFile =
        Path.Combine(NamesDirectory, "kozerdeku_csaladnev_2025.xlsx");

    private static readonly string FirstNamesFile =
        Path.Combine(NamesDirectory, "kozerdeku_utonevek_2025.xlsx");

    private static readonly string OutputFile =
        @"D:\Egyetem\7+félév\Szakdolgozat II\generalo algoritmusok\autótulaj generálás\jarmutulajdonosok.csv";

    private static readonly Random Random = new();

    // ============================================================
    // MAIN
    // ============================================================

    private static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;

        Console.WriteLine("Járműtulajdonos-generátor");
        Console.WriteLine("========================");
        Console.WriteLine();

        ValidateInputFiles();

        Console.WriteLine("Névstatisztikák betöltése...");

        List<WeightedName> lastNames = LoadLastNames(LastNamesFile);

        FirstNameData firstNames = LoadFirstNames(FirstNamesFile);

        Console.WriteLine($"Vezetéknevek: {lastNames.Count:N0}");
        Console.WriteLine($"Férfi keresztnevek: {firstNames.Male.Count:N0}");
        Console.WriteLine($"Női keresztnevek: {firstNames.Female.Count:N0}");
        Console.WriteLine();

        // Gyors súlyozott választáshoz kumulatív eloszlást készítünk.
        WeightedNameSelector lastNameSelector = new(lastNames);
        WeightedNameSelector maleNameSelector = new(firstNames.Male);
        WeightedNameSelector femaleNameSelector = new(firstNames.Female);

        Console.WriteLine("Járműszámok generálása...");

        int[] naturalVehicleCounts = GenerateNaturalPersonVehicleCounts(
            NaturalPersonCount,
            NaturalPersonVehicles);

        int[] legalVehicleCounts = GenerateLegalPersonVehicleCounts(
            LegalPersonCount,
            LegalPersonVehicles);

        // Összekeverjük, hogy ne az első rekordok kapják mindig
        // a több járművet.
        Shuffle(naturalVehicleCounts);
        Shuffle(legalVehicleCounts);

        Console.WriteLine("CSV generálása...");
        Console.WriteLine();

        GenerateCsv(
            lastNameSelector,
            maleNameSelector,
            femaleNameSelector,
            naturalVehicleCounts,
            legalVehicleCounts);

        Console.WriteLine();
        Console.WriteLine("Generálás kész.");
        Console.WriteLine($"CSV: {OutputFile}");
        Console.WriteLine();

        Console.WriteLine($"Tulajdonosok száma: {TotalOwners:N0}");
        Console.WriteLine($"  Természetes személy: {NaturalPersonCount:N0}");
        Console.WriteLine($"  Jogi személy:        {LegalPersonCount:N0}");
        Console.WriteLine();

        Console.WriteLine(
            $"Természetes személyek járművei: {naturalVehicleCounts.Sum():N0}");

        Console.WriteLine(
            $"Jogi személyek járművei:        {legalVehicleCounts.Sum():N0}");

        Console.WriteLine(
            $"Összes jármű:                    " +
            $"{naturalVehicleCounts.Sum() + legalVehicleCounts.Sum():N0}");
    }

    // ============================================================
    // FÁJLOK ELLENŐRZÉSE
    // ============================================================

    private static void ValidateInputFiles()
    {
        if (!File.Exists(LastNamesFile))
        {
            throw new FileNotFoundException(
                $"Nem található a családnév fájl:\n{LastNamesFile}");
        }

        if (!File.Exists(FirstNamesFile))
        {
            throw new FileNotFoundException(
                $"Nem található az utónév fájl:\n{FirstNamesFile}");
        }
    }

    // ============================================================
    // VEZETÉKNEVEK BETÖLTÉSE
    // ============================================================

    private static List<WeightedName> LoadLastNames(string path)
    {
        var result = new List<WeightedName>();

        using var workbook = new XLWorkbook(path);

        foreach (IXLWorksheet worksheet in workbook.Worksheets)
        {
            var rows = worksheet.RowsUsed().ToList();

            foreach (IXLRow row in rows)
            {
                string[] cells = row.CellsUsed()
                    .Select(c => c.GetFormattedString().Trim())
                    .ToArray();

                if (cells.Length < 2)
                    continue;

                // Megkeressük a sorban a nevet és az ahhoz
                // tartozó gyakorisági számot.
                string? name = null;
                long frequency = 0;

                for (int i = 0; i < cells.Length; i++)
                {
                    string value = cells[i];

                    if (string.IsNullOrWhiteSpace(value))
                        continue;

                    if (TryParseNumber(value, out long number))
                    {
                        if (number > frequency)
                            frequency = number;
                    }
                    else if (name == null && LooksLikeName(value))
                    {
                        name = NormalizeName(value);
                    }
                }

                if (!string.IsNullOrWhiteSpace(name) && frequency > 0)
                {
                    result.Add(new WeightedName(name, frequency));
                }
            }
        }

        result = MergeDuplicateNames(result);

        if (result.Count == 0)
        {
            throw new Exception(
                "Nem sikerült vezetékneveket beolvasni az Excel-fájlból.");
        }

        return result;
    }

    // ============================================================
    // KERESZTNEVEK BETÖLTÉSE
    // ============================================================

    private static FirstNameData LoadFirstNames(string path)
    {
        var male = new List<WeightedName>();
        var female = new List<WeightedName>();

        using var workbook = new XLWorkbook(path);

        foreach (IXLWorksheet worksheet in workbook.Worksheets)
        {
            string sheetName = worksheet.Name.ToLowerInvariant();

            bool maleSheet =
                sheetName.Contains("férfi") ||
                sheetName.Contains("ferfi") ||
                sheetName.Contains("male");

            bool femaleSheet =
                sheetName.Contains("nő") ||
                sheetName.Contains("no") ||
                sheetName.Contains("female");

            var rows = worksheet.RowsUsed().ToList();

            foreach (IXLRow row in rows)
            {
                var cells = row.CellsUsed()
                    .Select(c => c.GetFormattedString().Trim())
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    .ToList();

                // Megpróbálunk név + gyakoriság párokat találni.
                for (int i = 0; i < cells.Count - 1; i++)
                {
                    string possibleName = cells[i];
                    string possibleNumber = cells[i + 1];

                    if (!LooksLikeName(possibleName))
                        continue;

                    if (!TryParseNumber(possibleNumber, out long frequency))
                        continue;

                    if (frequency <= 0)
                        continue;

                    string name = NormalizeName(possibleName);

                    if (maleSheet)
                    {
                        male.Add(new WeightedName(name, frequency));
                    }
                    else if (femaleSheet)
                    {
                        female.Add(new WeightedName(name, frequency));
                    }
                }
            }
        }

        /*
         * Ha az Excel munkalapnevei alapján nem sikerült
         * elkülöníteni a férfi/női neveket, akkor megpróbáljuk
         * az oszlopfejlécek alapján.
         */
        if (male.Count == 0 || female.Count == 0)
        {
            ReadFirstNamesByHeaders(path, male, female);
        }

        male = MergeDuplicateNames(male);
        female = MergeDuplicateNames(female);

        if (male.Count == 0)
        {
            throw new Exception(
                "Nem sikerült férfi keresztneveket beolvasni.");
        }

        if (female.Count == 0)
        {
            throw new Exception(
                "Nem sikerült női keresztneveket beolvasni.");
        }

        return new FirstNameData(male, female);
    }

    private static void ReadFirstNamesByHeaders(
        string path,
        List<WeightedName> male,
        List<WeightedName> female)
    {
        using var workbook = new XLWorkbook(path);

        foreach (IXLWorksheet worksheet in workbook.Worksheets)
        {
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            for (int col = 1; col <= lastColumn; col++)
            {
                string header = worksheet.Cell(1, col)
                    .GetFormattedString()
                    .ToLowerInvariant();

                bool isMale =
                    header.Contains("férfi") ||
                    header.Contains("ferfi");

                bool isFemale =
                    header.Contains("nő") ||
                    header.Contains("noi") ||
                    header.Contains("női");

                if (!isMale && !isFemale)
                    continue;

                for (int row = 2; row <= lastRow; row++)
                {
                    string name = worksheet.Cell(row, col)
                        .GetFormattedString()
                        .Trim();

                    if (!LooksLikeName(name))
                        continue;

                    long frequency = 1;

                    // A következő oszlopot gyakoriságnak feltételezzük.
                    if (col + 1 <= lastColumn)
                    {
                        string value = worksheet.Cell(row, col + 1)
                            .GetFormattedString();

                        if (TryParseNumber(value, out long parsed) &&
                            parsed > 0)
                        {
                            frequency = parsed;
                        }
                    }

                    name = NormalizeName(name);

                    if (isMale)
                        male.Add(new WeightedName(name, frequency));

                    if (isFemale)
                        female.Add(new WeightedName(name, frequency));
                }
            }
        }
    }

    // ============================================================
    // CSV GENERÁLÁS
    // ============================================================

    private static void GenerateCsv(
        WeightedNameSelector lastNames,
        WeightedNameSelector maleNames,
        WeightedNameSelector femaleNames,
        int[] naturalVehicles,
        int[] legalVehicles)
    {
        Directory.CreateDirectory(
            Path.GetDirectoryName(OutputFile)!);

        using var writer = new StreamWriter(
            OutputFile,
            false,
            new UTF8Encoding(true),
            bufferSize: 1024 * 1024);

        writer.WriteLine(
            "id;vezeteknev;masodik_vezeteknev;" +
            "keresztnev;masodik_keresztnev;" +
            "jogi_szemely;jarmuvek_szama");

        long id = 1;

        // --------------------------------------------------------
        // TERMÉSZETES SZEMÉLYEK
        // --------------------------------------------------------

        for (int i = 0; i < NaturalPersonCount; i++)
        {
            bool male = Random.NextDouble() < 0.48;

            WeightedNameSelector firstNameSelector =
                male ? maleNames : femaleNames;

            string lastName = lastNames.Next();

            string secondLastName =
                Random.NextDouble() < SecondLastNameProbability
                    ? GetDifferentName(lastNames, lastName)
                    : "";

            string firstName = firstNameSelector.Next();

            string secondFirstName =
                Random.NextDouble() < SecondFirstNameProbability
                    ? GetDifferentName(firstNameSelector, firstName)
                    : "";

            WriteCsvRow(
                writer,
                id,
                lastName,
                secondLastName,
                firstName,
                secondFirstName,
                0,
                naturalVehicles[i]);

            id++;

            PrintProgress(id - 1);
        }

        // --------------------------------------------------------
        // JOGI SZEMÉLYEK
        // --------------------------------------------------------
        //
        // Mivel a kért rekordstruktúra minden rekordnál
        // vezeték- és keresztnevet tartalmaz, a jogi személyekhez
        // is generálunk egy személynevet.
        //
        // Ez értelmezhető például a céghez tartozó kapcsolattartó /
        // képviselő neveként.
        // --------------------------------------------------------

        for (int i = 0; i < LegalPersonCount; i++)
        {
            bool male = Random.NextDouble() < 0.48;

            WeightedNameSelector firstNameSelector =
                male ? maleNames : femaleNames;

            string lastName = lastNames.Next();

            string secondLastName =
                Random.NextDouble() < SecondLastNameProbability
                    ? GetDifferentName(lastNames, lastName)
                    : "";

            string firstName = firstNameSelector.Next();

            string secondFirstName =
                Random.NextDouble() < SecondFirstNameProbability
                    ? GetDifferentName(firstNameSelector, firstName)
                    : "";

            WriteCsvRow(
                writer,
                id,
                lastName,
                secondLastName,
                firstName,
                secondFirstName,
                1,
                legalVehicles[i]);

            id++;

            PrintProgress(id - 1);
        }
    }

    private static void WriteCsvRow(
        StreamWriter writer,
        long id,
        string lastName,
        string secondLastName,
        string firstName,
        string secondFirstName,
        int legalPerson,
        int vehicleCount)
    {
        writer.Write(id);
        writer.Write(';');

        writer.Write(EscapeCsv(lastName));
        writer.Write(';');

        writer.Write(EscapeCsv(secondLastName));
        writer.Write(';');

        writer.Write(EscapeCsv(firstName));
        writer.Write(';');

        writer.Write(EscapeCsv(secondFirstName));
        writer.Write(';');

        writer.Write(legalPerson);
        writer.Write(';');

        writer.WriteLine(vehicleCount);
    }

    // ============================================================
    // TERMÉSZETES SZEMÉLYEK JÁRMŰSZÁMA
    // ============================================================

    private static int[] GenerateNaturalPersonVehicleCounts(
        int ownerCount,
        int totalVehicles)
    {
        /*
         * Minden tulajdonosnak legalább 1 járműve van.
         *
         * 3 391 650 tulajdonos
         * 4 470 202 jármű
         *
         * Tehát 1 078 552 "plusz" járművet kell elosztanunk.
         *
         * A súlyozás célja:
         *
         * 1 jármű -> nagyon gyakori
         * 2 jármű -> gyakori
         * 3 jármű -> ritkább
         * 4+       -> egyre ritkább
         */

        int[] vehicles = Enumerable
            .Repeat(1, ownerCount)
            .ToArray();

        int remaining = totalVehicles - ownerCount;

        if (remaining < 0)
            throw new Exception(
                "Kevesebb jármű van, mint tulajdonos.");

        // Először véletlenszerűen osztjuk a plusz járműveket,
        // egyre csökkenő valószínűséggel.

        while (remaining > 0)
        {
            int index = Random.Next(ownerCount);

            int current = vehicles[index];

            double probability = current switch
            {
                1 => 1.00,
                2 => 0.30,
                3 => 0.10,
                4 => 0.04,
                5 => 0.015,
                _ => 0.005
            };

            if (Random.NextDouble() <= probability)
            {
                vehicles[index]++;
                remaining--;
            }
        }

        return vehicles;
    }

    // ============================================================
    // JOGI SZEMÉLYEK JÁRMŰSZÁMA
    // ============================================================

    private static int[] GenerateLegalPersonVehicleCounts(
        int ownerCount,
        int totalVehicles)
    {
        /*
         * 248 800 jogi személyhez összesen 870 634 jármű tartozik.
         *
         * Átlag:
         *
         * 870 634 / 248 800 ~= 3,5 jármű / jogi személy
         *
         * A céges eloszlásnál hosszabb farkat használunk:
         * sok 1-2 járműves cég mellett vannak nagyobb flották is.
         */

        int[] vehicles = Enumerable
            .Repeat(1, ownerCount)
            .ToArray();

        int remaining = totalVehicles - ownerCount;

        while (remaining > 0)
        {
            int index = Random.Next(ownerCount);

            int current = vehicles[index];

            double probability = current switch
            {
                1 => 1.00,
                2 => 0.90,
                3 => 0.75,
                4 => 0.60,
                5 => 0.45,
                6 => 0.35,
                7 => 0.28,
                8 => 0.22,
                9 => 0.18,

                <= 15 => 0.12,
                <= 25 => 0.07,
                <= 50 => 0.035,
                <= 100 => 0.015,

                _ => 0.004
            };

            if (Random.NextDouble() <= probability)
            {
                vehicles[index]++;
                remaining--;
            }
        }

        return vehicles;
    }

    // ============================================================
    // SEGÉDFÜGGVÉNYEK
    // ============================================================

    private static string GetDifferentName(
        WeightedNameSelector selector,
        string original)
    {
        for (int i = 0; i < 20; i++)
        {
            string result = selector.Next();

            if (!result.Equals(
                    original,
                    StringComparison.OrdinalIgnoreCase))
            {
                return result;
            }
        }

        return selector.Next();
    }

    private static void Shuffle(int[] array)
    {
        for (int i = array.Length - 1; i > 0; i--)
        {
            int j = Random.Next(i + 1);

            (array[i], array[j]) =
                (array[j], array[i]);
        }
    }

    private static string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "";

        if (value.Contains(';') ||
            value.Contains('"') ||
            value.Contains('\n') ||
            value.Contains('\r'))
        {
            return "\"" +
                   value.Replace("\"", "\"\"") +
                   "\"";
        }

        return value;
    }

    private static bool TryParseNumber(
        string value,
        out long number)
    {
        string cleaned = value
            .Replace(" ", "")
            .Replace("\u00A0", "")
            .Replace(".", "")
            .Replace(",", "");

        return long.TryParse(
            cleaned,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out number);
    }

    private static bool LooksLikeName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        if (value.Length < 2 || value.Length > 50)
            return false;

        string lower = value.ToLowerInvariant();

        string[] forbidden =
        {
            "név",
            "nev",
            "sorszám",
            "sorszam",
            "darab",
            "szám",
            "szam",
            "gyakoriság",
            "gyakorisag",
            "összes",
            "osszes",
            "férfi",
            "ferfi",
            "női",
            "noi",
            "helyezés",
            "helyezes"
        };

        if (forbidden.Any(x => lower.Contains(x)))
            return false;

        return value.Any(char.IsLetter);
    }

    private static string NormalizeName(string name)
    {
        name = name.Trim();

        if (name.All(c =>
                !char.IsLetter(c) ||
                char.IsUpper(c)))
        {
            TextInfo textInfo =
                CultureInfo.GetCultureInfo("hu-HU").TextInfo;

            name = textInfo.ToTitleCase(
                name.ToLower(
                    CultureInfo.GetCultureInfo("hu-HU")));
        }

        return name;
    }

    private static List<WeightedName> MergeDuplicateNames(
        List<WeightedName> source)
    {
        return source
            .GroupBy(
                x => x.Name,
                StringComparer.Create(
                    CultureInfo.GetCultureInfo("hu-HU"),
                    true))
            .Select(g =>
                new WeightedName(
                    g.First().Name,
                    g.Max(x => x.Weight)))
            .OrderByDescending(x => x.Weight)
            .ToList();
    }

    private static void PrintProgress(long generated)
    {
        if (generated % 100_000 != 0)
            return;

        double percent =
            generated / (double)TotalOwners * 100.0;

        Console.Write(
            $"\rGenerálva: {generated:N0} / " +
            $"{TotalOwners:N0} ({percent:F1}%)");
    }
}

// ================================================================
// SÚLYOZOTT NÉV
// ================================================================

internal sealed class WeightedName
{
    public string Name { get; }
    public long Weight { get; }

    public WeightedName(string name, long weight)
    {
        Name = name;
        Weight = weight;
    }
}

// ================================================================
// SÚLYOZOTT NÉVVÁLASZTÓ
// ================================================================
//
// Nem minden választáskor járjuk végig a teljes névlistát.
// Egyszer elkészítjük a kumulatív súlyokat, utána bináris
// kereséssel választunk.
//
// Ez 3,6 millió rekordnál jelentősen gyorsabb.
// ================================================================

internal sealed class WeightedNameSelector
{
    private readonly string[] _names;
    private readonly long[] _cumulativeWeights;
    private readonly long _totalWeight;

    public WeightedNameSelector(
        IReadOnlyList<WeightedName> names)
    {
        if (names.Count == 0)
            throw new ArgumentException(
                "A névlista nem lehet üres.");

        _names = new string[names.Count];
        _cumulativeWeights = new long[names.Count];

        long cumulative = 0;

        for (int i = 0; i < names.Count; i++)
        {
            _names[i] = names[i].Name;

            cumulative += Math.Max(1, names[i].Weight);

            _cumulativeWeights[i] = cumulative;
        }

        _totalWeight = cumulative;
    }

    public string Next()
    {
        long value = Random.Shared.NextInt64(
            1,
            _totalWeight + 1);

        int index = Array.BinarySearch(
            _cumulativeWeights,
            value);

        if (index < 0)
            index = ~index;

        return _names[index];
    }
}

// ================================================================
// KERESZTNÉV ADATOK
// ================================================================

internal sealed class FirstNameData
{
    public List<WeightedName> Male { get; }
    public List<WeightedName> Female { get; }

    public FirstNameData(
        List<WeightedName> male,
        List<WeightedName> female)
    {
        Male = male;
        Female = female;
    }
}