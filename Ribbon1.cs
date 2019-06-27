using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;

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
                    String LCANumber = Match.Groups[0].Value;
                    LCANumbers.Add(LCANumber);

                }
                // Opens explorer to navigate to folder
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    String FilePath = fbd.SelectedPath;

                    // Get the files in the directory
                    DirectoryInfo Dir = new DirectoryInfo(@FilePath);
                    FileInfo[] Files = Dir.GetFiles();

                    // Goes through each file in the selected folder
                    foreach (FileInfo CurrentFile in Files)
                    {
                        // Gets the file name and the file extension
                        String FileName = Path.GetFileNameWithoutExtension(CurrentFile.FullName);
                        String Ex = Path.GetExtension(CurrentFile.FullName);

                        // Goes through all of the selected emails 
                        foreach (String LCANumber in LCANumbers)
                        {
                            if(LCANumber.Contains(FileName) || LCANumber.Contains(Regex.Replace(FileName, @"[a-zA-Z]+$", "")))
                            {
                                if (Regex.IsMatch(FileName, @"[a-zA-Z]+$"))
                                {
                                    File.Move(CurrentFile.FullName, CurrentFile.FullName.Replace(CurrentFile.Name, $"{LCANumber} Signed Posting Attestation{Ex}"));
                                }
                                else
                                {
                                    File.Move(CurrentFile.FullName, CurrentFile.FullName.Replace(CurrentFile.Name, $"C {LCANumber}{Ex}"));
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }
    }
}
