using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

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

            // Pattern: <ix:nonNumeric ... id="FactXXXXX">selectedText</ix:nonNumeric>
            var regex = new System.Text.RegularExpressions.Regex(
                $@"(<ix:nonNumeric[^>]*id=""(Fact\d+)""[^>]*>)\s*{System.Text.RegularExpressions.Regex.Escape(selectedText)}\s*(</ix:nonNumeric>)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            );

            var match = regex.Match(html);
            if (!match.Success)
            {
                System.Windows.Forms.MessageBox.Show($"Could not find Fact tag for: {selectedText}", "Not Found");
                return;
            }

            string factId = match.Groups[2].Value;

            // Prompt for new value
            string newValue = Microsoft.VisualBasic.Interaction.InputBox(
                $"Enter new value for {factId} (was \"{selectedText}\"):", "Replace Fact", selectedText);

            if (string.IsNullOrEmpty(newValue) || newValue == selectedText)
                return;

            // Replace inner value
            string updatedHtml = regex.Replace(html, $"{match.Groups[1].Value}{newValue}{match.Groups[3].Value}");

            System.IO.File.WriteAllText(htmlPath, updatedHtml);

            Sel.Text = newValue;

            System.Windows.Forms.MessageBox.Show($"Updated {factId} to: {newValue}", "Fact Replaced");
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
