using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace LCARenamerAddIn
{
    public partial class TestRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        { 

        }

        private void TestButton_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (explorer != null)
            {
                List<String> LCANumbers = new List<String>();
                foreach (MailItem email in new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection)
                {
                    Match Match = Regex.Match(email.Subject, @"I[0-9\-]*$");
                    System.Windows.Forms.MessageBox.Show(Match.Groups[0].Value);
                    String LCANumber = Match.Groups[0].Value;
                    LCANumbers.Add(LCANumber);

                }
            }
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if(fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                MessageBox.Show(fbd.SelectedPath);
            }
        }
    }
}
