using Microsoft.Office.Interop.Word;
using w = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace highlight
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.buttonAnnotateFact.Enabled = true;
        }

        private void buttonAnnotateFact_Click(object sender, RibbonControlEventArgs e)
        {
            w.Selection selection = Globals.ThisAddIn.Application.Selection;
            string selectedText = selection.Text.Trim();

            if (!string.IsNullOrEmpty(selectedText))
            {
                System.Windows.Forms.MessageBox.Show("Selected Text (Fact): " + selectedText, "Fact Annotation");
                // Replace with your annotation form if needed
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please select some text to annotate.", "No Selection");
            }
        }

        private void toggleAnnotateFact_Click(object sender, RibbonControlEventArgs e)
        {
            // Optional: You could show status to user
            if (toggleAnnotateFact.Checked)
            {
                System.Windows.Forms.MessageBox.Show("Modify mode ON");
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Modify mode OFF");
            }
        }

        private void ReplaceDisclosureButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ReplaceDisclosure();
        }
    }
}
