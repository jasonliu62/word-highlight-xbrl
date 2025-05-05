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

            // Step 2: Read original HTML
            string html = File.ReadAllText(originalHtmlPath);

            // Match from Item... up to and including <p><b>SIGNATURE</b></p>
            var htmlRegex = new Regex(
                @"(<(p|td|span|div)[^>]*?>\s*<b>\s*Item\s.*?</b>.*?)(?=<p[^>]*?>\s*<b>SIGNATURE</b>\s*</p>)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);


            Match htmlMatch = htmlRegex.Match(html);
            if (!htmlMatch.Success)
            {
                System.Windows.Forms.MessageBox.Show("Could not locate disclosure in HTML.");
                return;
            }

            // Convert .docx to filtered HTML
            string tempHtmlPath = @"C:\Users\byung\Downloads\666\temp.html";
            Word.Document insertToHtml = wordApp.Documents.Open(insertDocxPath);
            insertToHtml.SaveAs2(tempHtmlPath, Word.WdSaveFormat.wdFormatFilteredHTML);
            insertToHtml.Close(false);

            string insertedHtml = File.ReadAllText(tempHtmlPath);

            // Extract only the main body content inside <div class=WordSection1>...</div>
            var mainBlockMatch = Regex.Match(insertedHtml, @"<div class=WordSection1>([\s\S]*?)</div>\s*</body>", RegexOptions.IgnoreCase);
            if (mainBlockMatch.Success)
            {
                insertedHtml = mainBlockMatch.Groups[1].Value;
            }

            // Normalize font
            for (int i = 0; i < 7; i++)
            {
                insertedHtml += "<p style=\"font: 10pt Times New Roman, Times, Serif; margin: 0pt 0\">&#160;</p>\n";
            }

            // Add a light horizontal rule
            insertedHtml += @"
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

            for (int i = 0; i < 3; i++)
            {
                insertedHtml += "<p style=\"font: 10pt Times New Roman, Times, Serif; margin: 0pt 0\">&#160;</p>\n";
            }


            // Final insertion: insertedHtml + preserved SIGNATURE block
            string updatedHtml = htmlRegex.Replace(html, insertedHtml, 1);
            File.WriteAllText(originalHtmlPath, updatedHtml);


            System.Windows.Forms.MessageBox.Show("Disclosure section replaced in both Word and HTML.");

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
