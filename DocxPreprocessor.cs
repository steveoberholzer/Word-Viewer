using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

namespace WordViewer;

/// <summary>
/// Applies in-memory fixups to a DOCX before handing it to Mammoth.
/// Currently handles:
///   - Missing w:lvl entries in abstractNum definitions (causes ArgumentOutOfRangeException)
///   - Dangling w:num → w:abstractNumId references
/// </summary>
internal static class DocxPreprocessor
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static MemoryStream Preprocess(string path)
    {
        // Copy the file into a MemoryStream so we can update the ZIP in place
        var ms = new MemoryStream();
        using (var fs = File.OpenRead(path))
            fs.CopyTo(ms);

        FixNumbering(ms);

        ms.Position = 0;
        return ms;
    }

    // ── Numbering fixup ───────────────────────────────────────────────────────

    private static void FixNumbering(MemoryStream ms)
    {
        const string entryPath = "word/numbering.xml";

        ms.Position = 0;
        using var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true);

        var entry = zip.GetEntry(entryPath);
        if (entry == null) return;          // document has no list definitions

        XDocument doc;
        using (var s = entry.Open())
            doc = XDocument.Load(s);

        bool changed = false;

        // 1. Ensure every abstractNum has levels 0–8 defined.
        //    Mammoth indexes directly by ilvl; a missing level causes AOORE.
        foreach (var abstractNum in doc.Descendants(W + "abstractNum"))
        {
            var defined = abstractNum
                .Elements(W + "lvl")
                .Select(l => (int?)l.Attribute(W + "ilvl") ?? 0)
                .ToHashSet();

            for (int i = 0; i <= 8; i++)
            {
                if (!defined.Contains(i))
                {
                    abstractNum.Add(MakeDefaultLevel(i));
                    changed = true;
                }
            }

            // Re-order levels by ilvl so Mammoth sees them in sequence
            if (changed)
            {
                var sorted = abstractNum.Elements(W + "lvl")
                    .OrderBy(l => (int?)l.Attribute(W + "ilvl") ?? 0)
                    .ToList();
                foreach (var lvl in sorted) lvl.Remove();
                foreach (var lvl in sorted) abstractNum.Add(lvl);
            }
        }

        // 2. Fix dangling num → abstractNumId references.
        var validAbstractIds = doc.Descendants(W + "abstractNum")
            .Select(a => (string?)a.Attribute(W + "abstractNumId"))
            .Where(id => id != null)
            .ToHashSet();

        var fallback = validAbstractIds.FirstOrDefault();

        foreach (var num in doc.Descendants(W + "num"))
        {
            var link = num.Element(W + "abstractNumId");
            var val  = (string?)link?.Attribute(W + "val");

            if (val != null && !validAbstractIds.Contains(val) && fallback != null && link != null)
            {
                link.SetAttributeValue(W + "val", fallback);
                changed = true;
            }
        }

        if (!changed) return;

        // Overwrite the entry with the patched XML
        entry.Delete();
        var newEntry = zip.CreateEntry(entryPath);
        using var writer = new StreamWriter(newEntry.Open(), System.Text.Encoding.UTF8);
        writer.Write(doc.Declaration != null
            ? doc.Declaration + "\n" + doc.Root
            : doc.ToString());
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static XElement MakeDefaultLevel(int index)
    {
        int indent  = 720 * (index + 1);
        int hanging = 360;

        return new XElement(W + "lvl",
            new XAttribute(W + "ilvl", index),
            new XElement(W + "start",   new XAttribute(W + "val", "1")),
            new XElement(W + "numFmt",  new XAttribute(W + "val", "bullet")),
            new XElement(W + "lvlText", new XAttribute(W + "val", "•")),
            new XElement(W + "lvlJc",   new XAttribute(W + "val", "left")),
            new XElement(W + "pPr",
                new XElement(W + "ind",
                    new XAttribute(W + "left",    indent.ToString()),
                    new XAttribute(W + "hanging", hanging.ToString()))));
    }
}
