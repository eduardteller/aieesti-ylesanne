using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

if (args.Length == 0)
{
    Console.WriteLine("Palun lisage dokument argumendina");
    return;
}

string filePath = args[0];

if (!File.Exists(filePath))
{
    Console.WriteLine("Viga: Fail ei eksisteeri!");
    return;
}

try
{
    using (var wordDoc = WordprocessingDocument.Open(filePath, true))
    {
        var eesti = new CultureInfo("et-EE");
        Body b = wordDoc.MainDocumentPart.Document.Body;

        KinnisvaraDokument dok = new KinnisvaraDokument(b);

        var location = dok.LeiaAsukoht();
        Teadmistebaas andmebaas = new Teadmistebaas(location);

        int loendur = 0;
        bool vordne = string.Compare(
            andmebaas.EestiKinnisvaraturg,
            dok.EestiKinnisvaraturg,
            ignoreCase: true,
            culture: eesti
        ) == 0;

        if (!vordne)
        {
            Console.WriteLine("Kinnisvaraturu andmed on uuendatud!");
            DokumentiRedaktor.UuendaDokumentiAndmed(
                b,
                "Eesti kinnisvaraturg",
                andmebaas.EestiKinnisvaraturg
            );

            loendur++;
        }

        vordne = string.Compare(
            andmebaas.MakromajanduslikTaust,
            dok.MakromajanduslikTaust,
            ignoreCase: true,
            culture: eesti
        ) == 0;

        if (!vordne)
        {
            Console.WriteLine("Makromajanduslik tausta andmed on uuendatud!");
            DokumentiRedaktor.UuendaDokumentiAndmed(
                b,
                "Makromajanduslik taust",
                andmebaas.MakromajanduslikTaust
            );

            loendur++;
        }

        vordne = string.Compare(
            andmebaas.Linnaosa,
            dok.Linnaosa,
            ignoreCase: true,
            culture: eesti
        ) == 0;

        if (!vordne)
        {

            if (location != "Õismäe" && location != "Kristiine")
            {
                throw new Exception("Tundmatu linnaosa");
            }

            Console.WriteLine($"{location} linnaosa korterituru tausta andmed on uuendatud!");

            DokumentiRedaktor.UuendaDokumentiAndmed(
                b,
                $"{location} linnaosa korteriturg",
                andmebaas.Linnaosa
            );

            loendur++;

        }

        if (loendur == 0)
        {
            Console.WriteLine("Andmed on ajakohased!");
            return;
        }

        wordDoc.MainDocumentPart.Document.Save();

        Console.WriteLine("\n" + "Dokument on salvestatud!");
    }

}
catch (Exception ex)
{
    Console.WriteLine($"Viga: {ex.Message}");
    return;
}
