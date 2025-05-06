using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Text.RegularExpressions;

namespace highlight
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            var ribbon = Globals.Ribbons.Ribbon1;

            if (!ribbon.toggleAnnotateFact.Checked)
                return;

            string selectedText = Sel.Text.Trim();
            if (string.IsNullOrEmpty(selectedText))
                return;

            string htmlPath = @"C:\Users\byung\Downloads\666\focus_8k.htm"; // Path

            string html = System.IO.File.ReadAllText(htmlPath);

            // Pattern: <ix:nonNumeric>selectedText</ix:nonNumeric>
            var regex = new System.Text.RegularExpressions.Regex(
                $@"(<ix:nonNumeric[^>]*>)\s*{System.Text.RegularExpressions.Regex.Escape(selectedText)}\s*(</ix:nonNumeric>)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            );

            var match = regex.Match(html);
            if (!match.Success)
            {
                System.Windows.Forms.MessageBox.Show($"Could not find Fact tag for: {selectedText}", "Not Found");
                return;
            }

            string openingTag = match.Groups[1].Value;
            string closingTag = match.Groups[2].Value;

            // Check if the tag is a date (based on format attribute)
            bool isDateFormatted = openingTag.Contains("format=\"ixt:datemonthdayyearen\"");

            string promptText = $"Enter new value for \"{selectedText}\":";
            if (isDateFormatted)
                promptText += "\n(e.g., January 21, 2019)";

            string newValue = Microsoft.VisualBasic.Interaction.InputBox(promptText, "Replace Fact", selectedText);

            if (string.IsNullOrEmpty(newValue) || newValue == selectedText)
                return;

            // Replace the ix:nonNumeric inner value
            string updatedHtml = regex.Replace(html, $"{openingTag}{newValue}{closingTag}");

            if (isDateFormatted && DateTime.TryParse(newValue, out DateTime parsedDate))
            {
                string isoDate = parsedDate.ToString("yyyy-MM-dd");

                updatedHtml = System.Text.RegularExpressions.Regex.Replace(
                    updatedHtml,
                    @"<xbrli:startDate>.*?</xbrli:startDate>",
                    $"<xbrli:startDate>{isoDate}</xbrli:startDate>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                );

                updatedHtml = System.Text.RegularExpressions.Regex.Replace(
                    updatedHtml,
                    @"<xbrli:endDate>.*?</xbrli:endDate>",
                    $"<xbrli:endDate>{isoDate}</xbrli:endDate>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                );
            }


            System.IO.File.WriteAllText(htmlPath, updatedHtml);

            Sel.Text = newValue;

            System.Windows.Forms.MessageBox.Show($"Updated fact to: {newValue}", "Fact Replaced");
        }


        public void ReplaceDisclosure()
        {
            string originalDocxPath = @"C:\Users\byung\Downloads\666\wat.docx"; // Your loaded Word file
            string originalHtmlPath = @"C:\Users\byung\Downloads\666\focus_8k.htm";     // Your 8-K HTML file
            string insertDocxPath = @"C:\Users\byung\Downloads\666\insertTest.docx";   // Your new content


            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document dc = app.Documents.Open(insertDocxPath, ReadOnly: false, Visible: true);

            foreach (Word.Table table in dc.Tables)
            {
                foreach (Word.Row row in table.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        Word.Range rng = cell.Range;
                        rng.ParagraphFormat.Reset(); 
                    }
                }
            }

            // Save and close the document
            dc.Save(); // Save over the original
            dc.Close();


            // Step 1: Load Word and remove disclosure block
            Word.Application wordApp = Globals.ThisAddIn.Application;
            Word.Document doc = wordApp.Documents.Open(originalDocxPath, ReadOnly: false);

            Word.Range disclosureRange = null;
            Word.Document insertDoc = null;
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                if (para.Range.Text.Trim().StartsWith("Item "))
                {
                    Word.Range start = para.Range;
                    Word.Range end = null;

                    foreach (Word.Paragraph p in doc.Paragraphs)
                    {
                        if (p.Range.Start > start.Start && p.Range.Text.Trim().Contains("SIGNATURE"))
                        {
                            end = p.Range;
                            break;
                        }
                    }

                    if (end != null)
                    {
                        int insertPos = start.Start; // Save where we will paste
                        disclosureRange = doc.Range(start.Start, end.Start);
                        disclosureRange.Delete();

                        // Step 2: Insert new content at the same position
                        Word.Range insertionPoint = doc.Range(insertPos, insertPos);
                        insertDoc = wordApp.Documents.Open(insertDocxPath);
                        insertDoc.Content.Copy();
                        insertDoc.Close(false);

                        insertionPoint.Paste();

                        Word.Range postInsert = doc.Range(insertionPoint.End, insertionPoint.End);
                        postInsert.InsertParagraphAfter(); // blank line
                        postInsert = doc.Range(postInsert.End, postInsert.End);
                        postInsert.InsertBreak(Word.WdBreakType.wdLineBreak); // force spacing
                        postInsert.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        postInsert.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth100pt;
                        postInsert.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorBlack;
                        postInsert.InsertParagraphAfter();

                        doc.Save();
                        doc.Close(false);

                        break;
                    }
                }
            }

            string html = File.ReadAllText(originalHtmlPath);

            // Match from first <b>Item X.XX</b> up to just before <p><b>SIGNATURE</b></p>
            var htmlRegex = new Regex(
                @"(<table[^>]*?>\s*<tr[^>]*?>\s*<td[^>]*?>\s*<span[^>]*?><b>\s*Item\s+\d+\.\d+.*?</b></span>.*?</table>[\s\S]*?)(?=<p[^>]*?><b>SIGNATURE</b></p>)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);


            Match htmlMatch = htmlRegex.Match(html);
            if (!htmlMatch.Success)
            {
                File.WriteAllText(@"C:\Users\byung\Downloads\666\DEBUG_failed_match.html", html);
                System.Windows.Forms.MessageBox.Show("Could not locate the first disclosure block before SIGNATURE.");
                return;
            }

            // Convert .docx to filtered HTML
            string tempHtmlPath = @"C:\Users\byung\Downloads\666\temp.html";
            Word.Document insertToHtml = wordApp.Documents.Open(insertDocxPath);
            insertToHtml.SaveAs2(tempHtmlPath, Word.WdSaveFormat.wdFormatFilteredHTML);
            insertToHtml.Close(false);

            // Fix encoding (Word exports in Windows-1252 / ANSI)
            string rawContent = File.ReadAllText(tempHtmlPath, Encoding.Default);
            File.WriteAllText(tempHtmlPath, rawContent, Encoding.UTF8);

            // Clean the filtered HTML
            string insertedRawHtml = File.ReadAllText(tempHtmlPath, Encoding.UTF8);
            string insertedHtml = CleanFilteredHtml(insertedRawHtml);
            File.WriteAllText(tempHtmlPath, insertedHtml, Encoding.UTF8);  // optional for inspection

            // Build visual spacing and divider
            string spacerBefore = string.Join("\n",
                Enumerable.Repeat("<p style=\"font: 10pt Times New Roman, Times, Serif; margin: 0pt 0\">&#160;</p>", 7));

            string divider = @"
<div style=""border-bottom: Black 1pt solid; margin-top: 6pt; margin-bottom: 6pt"">
  <table cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse; width: 100%; font-size: 10pt"">
    <tr style=""vertical-align: top; text-align: left"">
      <td style=""width: 33%"">&#160;</td>
      <td style=""width: 34%; text-align: center"">&#160;</td>
      <td style=""width: 33%; text-align: right"">&#160;</td>
    </tr>
  </table>
</div>
<div style=""break-before: page; margin-top: 6pt; margin-bottom: 6pt"">
  <p style=""margin: 0pt"">&#160;</p>
</div>";

            string spacerAfter = string.Join("\n",
                Enumerable.Repeat("<p style=\"font: 10pt Times New Roman, Times, Serif; margin: 0pt 0\">&#160;</p>", 3));

            // Final wrapper around inserted content
            string fullInsertBlock = $@"
<div style=""width:100%; font-family: 'Times New Roman', Times, Serif; font-size: 10pt;"">
{insertedHtml}
{spacerBefore}
{divider}
{spacerAfter}
</div>";

            // Optional debug: save what you're inserting
            File.WriteAllText(@"C:\Users\byung\Downloads\666\DEBUG_insert.html", fullInsertBlock);

            // Replace the original disclosure section
            string updatedHtml = html.Substring(0, htmlMatch.Index)
                      + fullInsertBlock
                      + html.Substring(htmlMatch.Index + htmlMatch.Length);

            // Save for verification
            File.WriteAllText(@"C:\Users\byung\Downloads\666\DEBUG_updated_output.html", updatedHtml, Encoding.UTF8);

            // Write to target
            File.WriteAllText(originalHtmlPath, updatedHtml, Encoding.UTF8);



            System.Windows.Forms.MessageBox.Show("Disclosure section replaced in both Word and HTML.");

        }

        private string CleanFilteredHtml(string rawHtml)
        {
            // 1. Remove <style> tags entirely
            rawHtml = Regex.Replace(rawHtml, @"<style[^>]*?>[\s\S]*?</style>", "", RegexOptions.IgnoreCase);

            // 2. Remove spans with visibility:hidden or display:none
            rawHtml = Regex.Replace(rawHtml, @"<span[^>]*?style\s*=\s*""[^""]*(visibility\s*:\s*hidden|display\s*:\s*none)[^""]*""[^>]*>.*?</span>", "", RegexOptions.IgnoreCase);

            // 3. Remove mso- prefixed styles from inline attributes
            rawHtml = Regex.Replace(rawHtml, @"style=""[^""]*?mso-[^""]*?""", "", RegexOptions.IgnoreCase);

            // 4. Remove align=center from <div> and replace with normal div
            rawHtml = Regex.Replace(rawHtml, @"<div\s+align\s*=\s*[""']?center[""']?\s*>", "<div>", RegexOptions.IgnoreCase);

            // 5. Fix all table styles: strip fixed width and inject responsive style
            rawHtml = Regex.Replace(
                rawHtml,
                @"<table([^>]*)style=""([^""]*)""([^>]*)>",
                match =>
                {
                    string beforeStyle = match.Groups[1].Value;
                    string styleContent = match.Groups[2].Value;
                    string afterStyle = match.Groups[3].Value;

                    // Remove fixed width
                    styleContent = Regex.Replace(styleContent, @"width\s*:[^;]+;?", "", RegexOptions.IgnoreCase);

                    // Clean whitespace
                    styleContent = Regex.Replace(styleContent, @"\s+", " ").Trim();

                    // Append proper table styling
                    string fixedStyle = $"style=\"{styleContent}; width:100%; border-collapse:collapse; table-layout:auto;\"";

                    return $"<table{beforeStyle} {fixedStyle}{afterStyle}>";
                },
                RegexOptions.IgnoreCase
            );

            // 6. Wrap each <table> in a scrollable <div>
            rawHtml = Regex.Replace(
                rawHtml,
                @"(<table[^>]*>[\s\S]*?</table>)",
                @"<div style=""overflow-x:auto; width:100%;"">$1</div>",
                RegexOptions.IgnoreCase
            );

            // 7. Extract content from WordSection1 and wrap it in a responsive container
            Match bodyMatch = Regex.Match(rawHtml, @"<div class=WordSection1>([\s\S]*?)</div>\s*</body>", RegexOptions.IgnoreCase);
            if (bodyMatch.Success)
            {
                string body = bodyMatch.Groups[1].Value;
                rawHtml = $@"
<div style='max-width:100%; padding-left:10.35%; padding-right:10.35%; position:relative;'>
{body}
</div>";
            }

            return rawHtml;
        }





        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
