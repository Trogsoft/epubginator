using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using Trogsoft.CommandLine;

namespace epubginator
{

    [Verb("paginate", HelpText = "Paginate an epub file.")]
    public class Paginator : Verb
    {
        private FileStream fs;
        private ZipArchive epub;

        private class ManifestEntry
        {
            public string Id { get; set; }
            public string Href { get; set; }
            public string MediaType { get; set; }
            public List<int> Pages { get; set; } = new List<int>();

            public void AddPage(int page)
            {
                Pages.Add(page);
            }

        }

        private class SpineEntry
        {
            public string Id { get; set; }
        }

        private class GuideEntry
        {
            public string Type { get; set; }
            public string Href { get; set; }
            public string Title { get; set; }
        }

        private class Document
        {
            public string Filename { get; set; }
            public HtmlDocument HtmlDocument { get; set; }
            public ZipArchiveEntry Entry { get; set; }
        }

        private List<Document> documents = new List<Document>();

        private Document GetHtmlDocument(string file)
        {

            if (documents.Any(x => x.Filename.Equals(file, StringComparison.CurrentCultureIgnoreCase)))
                return documents.SingleOrDefault(x => x.Filename.Equals(file, StringComparison.CurrentCultureIgnoreCase));

            var fileEntry = epub.Entries.SingleOrDefault(x => x.FullName == file);
            if (fileEntry == null)
                return null;

            using (var stream = fileEntry.Open())
            {
                MemoryStream ms = new MemoryStream();
                stream.CopyTo(ms);

                var html = Encoding.UTF8.GetString(ms.GetBuffer());
                HtmlDocument htmlDoc = new HtmlDocument();
                htmlDoc.OptionWriteEmptyNodes = true;
                htmlDoc.LoadHtml(html);

                var doc = new Document { Filename = file, HtmlDocument = htmlDoc, Entry = fileEntry };
                documents.Add(doc);

                return doc;
            }
        }
        private string GetTextContent(string file)
        {
            var htmlDoc = GetHtmlDocument(file);
            var body = htmlDoc.HtmlDocument.DocumentNode.SelectSingleNode("//body");
            return body.InnerText;
        }

        [Operation(isDefault: true)]
        [Parameter('f', "filename", HelpText = "The epub file to paginate.")]
        [Parameter('w', "wordsPerPage", HelpText = "How many words per page")]
        [Parameter("commit", HelpText = "Specify this parameter to write page information to the epub file.  Without it, only displays page information.")]
        public int Paginate(string filename, int wordsPerPage = 250, bool commit = false)
        {

            if (!File.Exists(filename))
            {
                Console.WriteLine("File not found.");
                return 1;
            }

            // backup
            var path = Path.GetDirectoryName(filename);
            var nf = Path.Combine(path, Path.GetFileNameWithoutExtension(filename) + ".paginated.epub");

            if (File.Exists(nf))
                File.Delete(nf);

            if (commit)
            {
                File.Copy(filename, nf);
                Console.WriteLine("Working file: " + nf);
            }
            else
            {
                nf = filename;
            }

            fs = new FileStream(nf, FileMode.Open, FileAccess.ReadWrite);
            epub = new ZipArchive(fs, ZipArchiveMode.Update);

            var containerEntry = epub.Entries.SingleOrDefault(x => x.FullName.Equals("meta-inf/container.xml", StringComparison.CurrentCultureIgnoreCase));
            if (containerEntry == null)
            {
                Console.WriteLine("Does not appear to be an epub file.");
                return 2;
            }

            MemoryStream containerStream = new MemoryStream();
            containerEntry.Open().CopyTo(containerStream);
            var containerText = Encoding.UTF8.GetString(containerStream.GetBuffer());

            var containerXml = new XmlDocument();
            containerXml.LoadXml(containerText.Trim('\0'));

            var rf = containerXml.GetElementsByTagName("rootfile");
            if (rf.Count == 0)
            {
                Console.WriteLine("No root file specified in container.xml");
                return 3;
            }

            var rootFile = rf[0].Attributes["full-path"];

            var opfEntry = epub.Entries.SingleOrDefault(x => x.FullName.Equals(rootFile.Value, StringComparison.CurrentCultureIgnoreCase));
            if (opfEntry == null)
            {
                Console.WriteLine("Could not find the OPF file.");
                return 4;
            }

            var opfStream = new MemoryStream();
            opfEntry.Open().CopyTo(opfStream);
            var opfText = Encoding.UTF8.GetString(opfStream.GetBuffer());

            var opf = new XmlDocument();
            XmlNamespaceManager nameSpaceManager = new XmlNamespaceManager(opf.NameTable);
            nameSpaceManager.AddNamespace("opf", "http://www.idpf.org/2007/opf");
            opf.LoadXml(opfText.Trim('\0'));

            List<ManifestEntry> manifest = new List<ManifestEntry>();
            List<SpineEntry> spine = new List<SpineEntry>();
            List<GuideEntry> guide = new List<GuideEntry>();

            foreach (XmlNode mf in opf.SelectNodes("/opf:package/opf:manifest/opf:item", nameSpaceManager))
            {
                var me = new ManifestEntry
                {
                    Href = mf.Attributes["href"].Value,
                    Id = mf.Attributes["id"].Value,
                    MediaType = mf.Attributes["media-type"].Value
                };
                manifest.Add(me);
            }

            foreach (XmlNode sf in opf.SelectNodes("/opf:package/opf:spine/opf:itemref", nameSpaceManager))
            {
                spine.Add(new SpineEntry
                {
                    Id = sf.Attributes["idref"].Value
                });
            }

            foreach (XmlNode gf in opf.SelectNodes("/opf:package/opf:guide/opf:reference", nameSpaceManager))
            {
                guide.Add(new GuideEntry
                {
                    Href = gf.Attributes["href"].Value,
                    Title = gf.Attributes["title"].Value,
                    Type = gf.Attributes["type"].Value,
                });
            }

            var wordCount = 0;
            var runningTotal = 0;
            var page = 1;
            foreach (var sf in spine)
            {
                bool modified = false;
                var me = manifest.SingleOrDefault(x => x.Id == sf.Id);
                if (me == null) continue;

                var doc = GetHtmlDocument(me.Href);
                var body = doc.HtmlDocument.DocumentNode.SelectSingleNode("//body");
                var nodes = body.Descendants().Where(x => new string[] { "p" }.Contains(x.Name));

                foreach (HtmlNode node in nodes)
                {
                    if (node.InnerText.Length < 2) continue;
                    wordCount += node.InnerText.Split(' ').Count();
                    runningTotal += node.InnerText.Split(' ').Count();

                    if (commit && runningTotal > wordsPerPage)
                    {
                        modified = true;
                        runningTotal -= wordsPerPage;
                        var pageBreak = doc.HtmlDocument.CreateElement("span");
                        pageBreak.Attributes.Add("role", "doc-pagebreak");
                        pageBreak.Attributes.Add("id", "pg" + page);
                        pageBreak.Attributes.Add("aria-label", page.ToString());
                        me.AddPage(page);
                        page++;

                        node.AppendChild(pageBreak);
                    }
                }

                if (modified)
                {
                    doc.Entry.Delete();
                    var e = epub.CreateEntry(doc.Filename, CompressionLevel.Optimal);
                    using (var sw = new StreamWriter(e.Open()))
                    {
                        sw.WriteLine(doc.HtmlDocument.DocumentNode.OuterHtml);
                    }
                }

            }

            if (commit)
            {

                // ncx
                var ncx = epub.Entries.SingleOrDefault(x => x.FullName.EndsWith(".ncx", StringComparison.CurrentCultureIgnoreCase));
                if (ncx != null)
                {

                    var ncxFilename = ncx.Name;
                    MemoryStream ms = new MemoryStream();
                    using (var str = ncx.Open())
                    {
                        str.CopyTo(ms);
                    }
                    var ncxmlString = Encoding.UTF8.GetString(ms.GetBuffer());

                    XmlDocument ncxml = new XmlDocument();
                    XmlNamespaceManager nsm = new XmlNamespaceManager(ncxml.NameTable);
                    nsm.AddNamespace("ncx", "http://www.daisy.org/z3986/2005/ncx/");
                    ncxml.LoadXml(ncxmlString.Trim('\0'));

                    var head = ncxml.SelectSingleNode("//ncx:head", nsm);

                    var foundMaxPage = false;
                    var foundTotalPage = false;
                    foreach (XmlNode meta in head.SelectNodes("ncx:meta", nsm))
                    {
                        if (meta.Attributes["name"].Value == "dtb:totalPageCount")
                        {
                            foundTotalPage = true;
                            meta.Attributes["content"].Value = (page - 1).ToString();
                        }
                        if (meta.Attributes["name"].Value == "dtb:maxPageNumber")
                        {
                            foundMaxPage = true;
                            meta.Attributes["content"].Value = (page - 1).ToString();
                        }
                    }

                    if (!foundTotalPage)
                    {
                        var newMeta = ncxml.CreateElement("meta", "http://www.daisy.org/z3986/2005/ncx/");
                        newMeta.SetAttribute("name", "dtb:totalPageCount");
                        newMeta.SetAttribute("content", (page - 1).ToString());
                        head.AppendChild(newMeta);
                    }

                    if (!foundMaxPage)
                    {
                        var newMeta = ncxml.CreateElement("meta", "http://www.daisy.org/z3986/2005/ncx/");
                        newMeta.SetAttribute("name", "dtb:maxPageNumber");
                        newMeta.SetAttribute("content", (page - 1).ToString());
                        head.AppendChild(newMeta);
                    }

                    var pageList = ncxml.SelectSingleNode("//ncx:pageList", nsm);
                    if (pageList != null)
                    {
                        pageList.RemoveAll();
                    }
                    else
                    {
                        pageList = ncxml.CreateElement("pageList");
                    }

                    var navLabel = ncxml.CreateElement("navLabel");
                    var navLabelText = ncxml.CreateElement("text");
                    navLabelText.InnerText = "Pages";
                    navLabel.AppendChild(navLabelText);
                    pageList.AppendChild(navLabel);

                    for (var p = 1; p < page; p++)
                    {

                        var pageX = ncxml.CreateElement("pageTarget");
                        pageX.SetAttribute("type", "normal");
                        pageX.SetAttribute("id", "pg" + p);
                        pageX.SetAttribute("value", p.ToString());
                        pageX.SetAttribute("playOrder", p.ToString());

                        var pnl = ncxml.CreateElement("navLabel");
                        var pnlt = ncxml.CreateElement("text");
                        pnlt.InnerText = p.ToString();
                        pnl.AppendChild(pnlt);
                        pageX.AppendChild(pnl);

                        var entity = manifest.SingleOrDefault(x => x.Pages.Contains(p));

                        var c = ncxml.CreateElement("content");
                        c.SetAttribute("src", entity.Href + "#pg" + p);

                        pageX.AppendChild(c);
                        pageList.AppendChild(pageX);

                    }

                    var root = ncxml.SelectSingleNode("//ncx:ncx", nsm);
                    root.AppendChild(pageList);

                    ncx.Delete();
                    var newncx = epub.CreateEntry(ncxFilename);
                    using (var sw = new StreamWriter(newncx.Open()))
                    {
                        sw.WriteLine(ncxml.DocumentElement.OuterXml);
                    }

                }

                // nav.xhtml 
                // todo: probably need to consult the spine to find this

                var nav = epub.Entries.SingleOrDefault(x => x.FullName.Equals("nav.xhtml", StringComparison.CurrentCultureIgnoreCase));
                if (nav != null)
                {
                    var navFilename = nav.Name;
                    MemoryStream ms = new MemoryStream();
                    using (var str = nav.Open())
                    {
                        str.CopyTo(ms);
                    }
                    var navString = Encoding.UTF8.GetString(ms.GetBuffer());
                    HtmlDocument hd = new HtmlDocument();

                    hd.OptionWriteEmptyNodes = true;
                    hd.LoadHtml(navString);

                    var body = hd.DocumentNode.SelectSingleNode("//body");
                    HtmlNode pln = null;
                    foreach (HtmlNode node in body.SelectNodes("//nav"))
                    {
                        if ((node.Attributes["epub:type"]?.Value ?? "") == "page-list")
                        {
                            pln = node;
                            break;
                        }
                    }

                    if (pln != null)
                        pln.Remove();

                    pln = hd.CreateElement("nav");
                    pln.Attributes.Add("epub:type", "page-list");
                    pln.Attributes.Add("hidden", "hidden");

                    body.AppendChild(pln);

                    var ol = hd.CreateElement("ol");

                    for (var p = 1; p < page; p++)
                    {
                        var entity = manifest.SingleOrDefault(x => x.Pages.Contains(p));
                        var li = hd.CreateElement("li");
                        var a = hd.CreateElement("a");
                        a.Attributes.Add("href", entity.Href + "#pg" + p);
                        var t = hd.CreateTextNode(p.ToString());
                        a.AppendChild(t);
                        li.AppendChild(a);
                        ol.AppendChild(li);
                    }

                    pln.AppendChild(ol);

                    nav.Delete();
                    var newnav = epub.CreateEntry(navFilename);
                    using (var sw = new StreamWriter(newnav.Open()))
                    {
                        sw.WriteLine(hd.DocumentNode.OuterHtml);
                    }

                }

            }

            Console.WriteLine("Words: " + wordCount);
            Console.WriteLine("Pages: " + wordCount / wordsPerPage);

            // save the page info.
            epub.Dispose();

            return 0;

        }
    }
}
