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

        var nimi = Path.GetFileName(filePath);
        var eesti = new CultureInfo("et-EE");
        Body b = wordDoc.MainDocumentPart.Document.Body;

        // Kui dokument on tühi siis ei ole mõtet jätkata
        if (b == null)
        {
            Console.WriteLine("Viga: Dokument on tühi!");
            return;
        }

        // initsialiseerin KinnisvaraDokument klassi ja Teadmistebaas klassi
        KinnisvaraDokument dok = new KinnisvaraDokument(b);
        Teadmistebaas andmebaas = new Teadmistebaas(dok.LeiaAsukoht());


        // Võrdlen dokumenti ja teadmistebaasi piirkondlikku ylevaade andmeid, kui teadmistebaaasi andmed on vananenud siis uuendan
        var PiirkiondlikYlevaadeAastaTeadmistebaasis = DocxTooristad.AastaArvTekstist(andmebaas.PiirkondlikYlevaade);
        var PiirkiondlikYlevaadeAastaDokumendis = DocxTooristad.AastaArvTekstist(dok.PiirkondlikYlevaade);

        if (PiirkiondlikYlevaadeAastaTeadmistebaasis != null && PiirkiondlikYlevaadeAastaDokumendis != null)
        {
            // Kui aasta on sama siis on andmed ajakohased
            if (PiirkiondlikYlevaadeAastaTeadmistebaasis == PiirkiondlikYlevaadeAastaDokumendis)
            {
                Console.WriteLine("Teadmistebaasi andmed on ajakohased!");
            }

            else if (PiirkiondlikYlevaadeAastaTeadmistebaasis < PiirkiondlikYlevaadeAastaDokumendis)
            {
                Console.WriteLine($"Teadmistebaasi andmed on vananenud! {PiirkiondlikYlevaadeAastaTeadmistebaasis} vs {PiirkiondlikYlevaadeAastaDokumendis}");

                var lok = dok.LeiaAsukoht();

                if (lok != "Õismäe" && lok != "Kristiine")
                {
                    throw new Exception("Tundmatu linnaosa");
                }

                andmebaas.UuendaPiirkondlikYlevaade(
                    dok.PiirkondlikYlevaade
                );

                Console.WriteLine($"{andmebaas.teadmistebaasiFailideNimed[2]} <- {nimi}\n");
            }
        }
        else
        {
            Console.WriteLine("WARN: Probleem Piirkondlike Ülevaatuste võrdlemisel. Jätan vahele...");
        }

        // Loenduriga saan näha kas hindamisfaili andmed olid ajakohased või mitte
        int loendur = 0;


        // Võrdlen kinnisvaraturu andmeid, kui need ei ole võrdsed siis uuendan
        bool vordne = string.Compare(
            andmebaas.EestiKinnisvaraturg,
            dok.EestiKinnisvaraturg,
            ignoreCase: true,
            culture: eesti
        ) == 0;

        if (!vordne)
        {
            Console.WriteLine("Kinnisvaraturu andmed on uuendatud!");
            Console.WriteLine($"{nimi} <- {andmebaas.teadmistebaasiFailideNimed[0]}\n");
            DokumentiRedaktor.UuendaDokumentiAndmed(
                b,
                "Eesti kinnisvaraturg",
                andmebaas.EestiKinnisvaraturg
            );

            loendur++;
        }

        // Võrdlen Makromajandusliku tausta andmed, kui need ei ole võrdsed siis uuendan
        vordne = string.Compare(
            andmebaas.MakromajanduslikTaust,
            dok.MakromajanduslikTaust,
            ignoreCase: true,
            culture: eesti
        ) == 0;

        if (!vordne)
        {
            Console.WriteLine("Makromajandusliku tausta andmed on uuendatud!");
            Console.WriteLine($"{nimi} <- {andmebaas.teadmistebaasiFailideNimed[1]}\n");
            DokumentiRedaktor.UuendaDokumentiAndmed(
                b,
                "Makromajanduslik taust",
                andmebaas.MakromajanduslikTaust
            );

            loendur++;
        }

        // Võrdlen piirkondliku ülevaate andmed, kui need ei ole võrdsed siis uuendan
        vordne = string.Compare(
            andmebaas.PiirkondlikYlevaade,
            dok.PiirkondlikYlevaade,
            ignoreCase: true,
            culture: eesti
        ) == 0;

        if (!vordne)
        {

            var lok = dok.LeiaAsukoht();

            if (lok != "Õismäe" && lok != "Kristiine")
            {
                throw new Exception("Tundmatu linnaosa");
            }

            Console.WriteLine($"{lok} linnaosa korterituru tausta andmed on uuendatud!");
            Console.WriteLine($"{nimi} <- {andmebaas.teadmistebaasiFailideNimed[2]}\n");

            DokumentiRedaktor.UuendaDokumentiAndmed(
                b,
                $"{lok} linnaosa korteriturg",
                andmebaas.PiirkondlikYlevaade
            );

            loendur++;

        }

        if (loendur == 0)
        {
            Console.WriteLine("Andmed on ajakohased!");
            return;
        }

        // Salvestan muudetud dokumendi
        wordDoc.MainDocumentPart.Document.Save();

        Console.WriteLine("\n" + "Dokument on salvestatud!");
    }

}
catch (Exception ex)
{
    Console.WriteLine($"Viga: {ex.Message}");
    return;
}
