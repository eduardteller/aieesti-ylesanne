using DocumentFormat.OpenXml.Wordprocessing;

public class KinnisvaraDokument
{
    public string HinnatavVara { get; private set; } = string.Empty;
    public string Kokkuvote { get; private set; } = string.Empty;
    public string MakromajanduslikTaust { get; private set; } = string.Empty;
    public string EestiKinnisvaraturg { get; private set; } = string.Empty;
    public string PiirkondlikYlevaade { get; private set; } = string.Empty;
    public string Turuvaartus { get; private set; } = string.Empty;

    public KinnisvaraDokument(Body b)
    {

        if (b == null)
        {
            throw new ArgumentNullException(nameof(b), "Dokument on tühi");
        }

        ParsiDokument(b);
    }

    private void ParsiDokument(Body b)
    {
        var pealkirjadeKaart = new Dictionary<string, Action<string>>(
          StringComparer.OrdinalIgnoreCase
      )
        {
            { "Turuülevaade", value => { } },
            { "Hinnatav vara:", value => HinnatavVara = value?.Trim() ?? string.Empty },
            { "Kokkuvõte:", value => Kokkuvote = value?.Trim() ?? string.Empty },
            {
                "Makromajanduslik taust",
                value => MakromajanduslikTaust = value?.Trim() ?? string.Empty
            },
            {
                "Eesti kinnisvaraturg",
                value => EestiKinnisvaraturg = value?.Trim() ?? string.Empty
            },
            { "Kristiine linnaosa korteriturg", value => PiirkondlikYlevaade = value?.Trim() ?? string.Empty },
            { "Õismäe linnaosa korteriturg", value => PiirkondlikYlevaade = value?.Trim() ?? string.Empty },
            { "Turuväärtus:", value => Turuvaartus = value?.Trim() ?? string.Empty }
        };

        try
        {
            string? peagunePealkiri = null;
            var peaguneSisu = new List<string>();

            foreach (var p in b.Elements<Paragraph>())
            {
                var text = DocxTooristad.LeiaTekst(p);

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (pealkirjadeKaart.ContainsKey(text))
                {
                    TootleSektsiooni(peagunePealkiri, peaguneSisu, pealkirjadeKaart);
                    peagunePealkiri = text;
                    peaguneSisu.Clear();
                }
                else if (peagunePealkiri != null)
                {
                    peaguneSisu.Add(text);
                }
            }

            TootleSektsiooni(peagunePealkiri, peaguneSisu, pealkirjadeKaart);
        }
        catch (Exception ex)
        {
            throw new Exception("Viga", ex);
        }
    }

    private void TootleSektsiooni(
        string? pealkiri,
        List<string> sisu,
        Dictionary<string, Action<string>> kaart
    )
    {
        if (pealkiri != null && sisu.Count > 0)
        {
            try
            {
                kaart[pealkiri](string.Join("\n", sisu));
            }
            catch (Exception ex)
            {
                throw new Exception(
                    $"Viga teksti tootlemisel '{pealkiri}'",
                    ex
                );
            }
        }
    }

    public string LeiaAsukoht()
    {

        if (string.IsNullOrEmpty(HinnatavVara))
        {
            throw new Exception(
                $"Dokumenti Struktuur ei Vasta Standardile (Hinnatav Vara): {HinnatavVara}"
            );
        }

        if (HinnatavVara.Contains("õismäe", StringComparison.OrdinalIgnoreCase))
        {
            return "Õismäe";
        }
        else if (HinnatavVara.Contains("kristiine", StringComparison.OrdinalIgnoreCase))
        {
            return "Kristiine";
        }

        throw new Exception(
            $"Viga Lokatsiooni Leidmisel: {HinnatavVara}"
        ); ;
    }

}
