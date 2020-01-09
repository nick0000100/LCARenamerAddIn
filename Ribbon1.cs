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
                FolderBrowserDialog FBD = new FolderBrowserDialog
                {
                    ShowNewFolderButton = false,
                    Description = "Select folder with LCAs"
                };
                if (FBD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    String FilePath = FBD.SelectedPath;

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
                            if (LCANumber.Contains(FileName) || LCANumber.Contains(Regex.Replace(FileName, @"[a-zA-Z]+$", "")) || LCANumber.Contains(Regex.Replace(FileName, @"\s[0-9a-zA-Z]*$", "")))
                            {
                                int Count = 1;
                                String NewFileName = "";

                                // Looks for blanket LCA
                                if(Regex.IsMatch(FileName, @"\s[0-9]+"))
                                {
                                    String BlanketNumber = Regex.Replace(FileName.Split(' ').Last(), @"[a-zA-Z]+$", "");

                                    // Appends 0s to blanket number if needed
                                    if (BlanketNumber.Length < 3)
                                    {
                                        String Zeroes = "";
                                        for(int i = 0; i < (3 - BlanketNumber.Length); i++)
                                        {
                                            Zeroes += "0";
                                        }
                                        BlanketNumber = Zeroes + BlanketNumber;
                                    }

                                    NewFileName = CurrentFile.FullName.Replace(CurrentFile.Name, $"FY21 Blanket [{BlanketNumber}] - [{LCANumber}]");

                                    // Blanket LCA Attestation
                                    if(Regex.IsMatch(FileName, @"[a-zA-Z]+$"))
                                    {
                                        NewFileName += $" Signed Posting Attestation";
                                    }

                                    NewFileName = Path.ChangeExtension(NewFileName, Ex);

                                }
                                // Normal LCA attestation
                                else if (Regex.IsMatch(FileName, @"[a-zA-Z]+$"))
                                {
                                    NewFileName = CurrentFile.FullName.Replace(CurrentFile.Name, $"{LCANumber} Signed Posting Attestation{Ex}");

                                    while(File.Exists(NewFileName))
                                    {
                                        NewFileName = CurrentFile.FullName.Replace(CurrentFile.Name, $"{LCANumber} Signed Posting Attestation ({++Count}){Ex}");
                                    }
                                }
                                // Normal LCA
                                else
                                {
                                    NewFileName = CurrentFile.FullName.Replace(CurrentFile.Name, $"C {LCANumber}{Ex}");

                                    while (File.Exists(NewFileName))
                                    {
                                        NewFileName = CurrentFile.FullName.Replace(CurrentFile.Name, $"C {LCANumber} ({++Count}){Ex}");
                                    }
                                }
                                File.Move(CurrentFile.FullName, NewFileName);
                                break;
                            }
                        }
                    }
                }
            }
        }
    }
}
