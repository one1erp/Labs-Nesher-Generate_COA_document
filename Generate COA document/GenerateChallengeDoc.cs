using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using DAL;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using DataTable = System.Data.DataTable;
using System.Diagnostics;
using Application = Microsoft.Office.Interop.Word.Application;

namespace Generate_COA_document
{
    public class GenerateChallengeDoc
    {
        private IDataLayer dal;
        private string _savedPath;
        private static int _typeForRow;
        public static string GetTypeOfChllenge;
        private const string EP = "ep";
        private const string USP = "USP";
        private const string USP_SPECIAL = "challenge";
        private string _coaName;

        public GenerateChallengeDoc(Sdg sdg, IDataLayer dal, string coaName)
        {
            this.dal = dal;
            _coaName = coaName;
            WordDoc(sdg);
        }

        private static Sample currentSample;
        public void WordDoc(Sdg sdg)
        {
            Object oMissing = Missing.Value;
            var wordApp = new Application();
            var wd = new Document();
            currentSample = sdg.Samples.FirstOrDefault(x => x.Status != "X");
            foreach (var aliq in currentSample.Aliqouts.Where(x => x.Status != "X").OrderByDescending(x => x.Children.Count))
            {
                int Col;

                if (aliq.Children.Count > 0)
                {

                    try
                    {
                        bool isGmp = false;
                        if (aliq.TestTemplateEx.IsGMP == "T") isGmp = true;

                        Object oTemplatePath = GetTemplatePath(aliq.TestTemplateEx.Workflow.Name, isGmp);//aliq.TestTemplateEx.IsGMP

                        wd = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
                    }
                    catch (Exception exception)
                    {
                        wd.Close();
                        wordApp.Application.Quit();
                        Common.Logger.WriteLogFile(exception);
                        throw;
                    }
                    TwoFirsetTable(sdg, wd);
                }
                Col = GetCol(aliq);
                if (Col != 0)
                {
                    FillTests(wd, aliq, Col);
                }
            }
            PermanentFields(wd, wordApp);
            SaveFile(wordApp, wd);
        }

        private static int GetCol(Aliquot aliq)
        {

            int Col = 0;
            if (aliq.Name.Contains("time 0"))
            {
                Col = 2;
            }
            if (aliq.Name.Contains("CHL 6h"))
            {
                Col = 3;
            }
            if (aliq.Name.Contains("CHL 24h"))
            {
                Col = 4;
            }
            if (aliq.Name.Contains("CHL 48h"))
            {
                Col = 5;
            }
            if (aliq.Name.Contains("CHL 7 days"))
            {
                if (GetTypeOfChllenge == USP)
                {
                    Col = 3;
                }
                if (GetTypeOfChllenge == USP_SPECIAL || GetTypeOfChllenge == EP)
                {
                    Col = 6;
                }
            }
            if (aliq.Name.Contains("CHL 14 days"))
            {
                if (GetTypeOfChllenge == USP)
                {
                    Col = 4;
                }
                if (GetTypeOfChllenge == USP_SPECIAL || GetTypeOfChllenge == EP)
                {
                    Col = 7;
                }
            }
            if (aliq.Name.Contains("CHL 28 days"))
            {
                if (GetTypeOfChllenge == USP)
                {
                    Col = 5;
                }
                if (GetTypeOfChllenge == USP_SPECIAL)
                {
                    Col = 9;
                }
                if (GetTypeOfChllenge == EP)
                {
                    Col = 8;
                }
            }
            if (aliq.Name.Contains("CHL 21 days"))
            {
                Col = 8;
            }
            return Col;
        }




        private static void FillTests(Document wd, Aliquot aliq, int Col)
        {
            foreach (var test in aliq.Tests)
            {

                bool uninoculated = false;
                var row = 0;
                switch (test.NAME)
                {
                    case "E. Coli":
                        row = 4;
                        break;
                    case "S.aureus":
                        row = 2;
                        break;
                    case "Ps. aeruginosa":
                        row = 3;
                        break;
                    case "Cd. albicans":
                        if (GetTypeOfChllenge == EP)
                        {
                            row = 4;
                        }
                        else
                        {
                            row = 5;
                        }

                        break;
                    case "A. Braziliensis":
                        if (GetTypeOfChllenge == EP)
                        {
                            row = 5;
                        }
                        else
                        {
                            row = 6;
                        }
                        break;
                    case "Uninoculated Control":
                        if (GetTypeOfChllenge == EP)
                        {
                            row = 6;
                        }
                        else if (GetTypeOfChllenge == USP)
                        {
                            row = 7;
                        }
                        else
                        {
                            row = 8;
                        }
                        uninoculated = true;
                        break;
                    default:
                        row = 7;
                        break;
                }

                ;
                if (row != 0)
                {
                    if (test.NAME != " ")
                    {
                        if (test.Results.FirstOrDefault() != null)
                        {
                            string finalResult = "";
                            var aaa = test.Results.Where(x => x.Name == ("Final Result")).FirstOrDefault();
                            if (aaa != null)
                            {
                                finalResult =
                                    test.Results.Where(x => x.Name == ("Final Result")).FirstOrDefault().
                                        FormattedResult;
                            }
                            if (uninoculated)
                            {
                                finalResult = test.Results.FirstOrDefault().FormattedResult;
                            }
                            if (Col == 2) // 0זמן
                            {
                                finalResult = test.Results.FirstOrDefault().FormattedResult;
                            }
                            if (row == 7 && GetTypeOfChllenge == USP_SPECIAL)
                            {
                                InsertVal(wd, 1, row, test.NAME);
                            }
                            InsertVal(wd, Col, row, finalResult);
                        }
                    }

                }
            }
            return;
        }

        private static void InsertVal(Document wd, int Col, int rowNum, string val)
        {
            string desc = "aa";
            // MessageBox.Show(Col + " " + rowNum);
            var row = wd.Tables[2].Rows[rowNum];
            //  MessageBox.Show("1");
            row.Cells[Col].Range.Text = numToExponent(val, desc);
            //  MessageBox.Show("2");
        }

        private static void TwoFirsetTable(Sdg sdg, Document wd)
        {
            Row row = wd.Tables[1].Rows[1];
            if (sdg.Client != null) row.Cells[2].Range.Text = sdg.Client.Name;

            row = wd.Tables[1].Rows[2];
            if (sdg.ContactName != null) row.Cells[2].Range.Text = sdg.ContactName;

            row = wd.Tables[1].Rows[3];
            if (sdg.Address != null) row.Cells[2].Range.Text = sdg.Address;

            row = wd.Tables[1].Rows[4];
            if (currentSample.Description != null) row.Cells[2].Range.Text = currentSample.Description;

            row = wd.Tables[1].Rows[1];
            if (sdg.Phone != null) row.Cells[5].Range.Text = sdg.Phone;

            row = wd.Tables[1].Rows[2];


            string email = "";
            if (!String.IsNullOrEmpty(sdg.Emai))
                email = sdg.Emai.Split(';')[0];
            row.Cells[5].Range.Text = email;

            row = wd.Tables[1].Rows[3];
            if (sdg.SampledByOperator != null) row.Cells[5].Range.Text = sdg.SampledByOperator.Name;

            row = wd.Tables[1].Rows[4];
            row.Cells[5].Range.Text = currentSample.Batch;

            row = wd.Tables[1].Rows[1];
            row.Cells[8].Range.Text = Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yyyy");

            row = wd.Tables[1].Rows[2];
            row.Cells[8].Range.Text = Convert.ToDateTime(sdg.CREATED_ON).ToString("dd/MM/yyyy");

            row = wd.Tables[1].Rows[3];
            row.Cells[8].Range.Text = sdg.ExternalReference;

            row = wd.Tables[1].Rows[4];
            row.Cells[8].Range.Text = sdg.Name;

            row = wd.Tables[1].Rows[5];
            row.Cells[2].Range.Text = sdg.U_COA_REMARKS;

            var firstOrDefault = currentSample.Aliqouts.FirstOrDefault(a => a.Children.Count() > 0 && a.Status != "X");

            if (firstOrDefault != null)
            {
                var conclusion = firstOrDefault.Conclusion;

                row = wd.Tables[2].Rows[4 + _typeForRow];
                if (conclusion == "O")
                {
                    row.Cells[2].Range.Text = "The product is not being preserved in a satisfactory manner and does not meet the requirements of the " + firstOrDefault.TestTemplateEx.Standard;
                }
                else if (conclusion == "I")
                {
                    row.Cells[2].Range.Text = "The product  is being preserved in a satisfactory manner and meets the requirements of the " + firstOrDefault.TestTemplateEx.Standard;
                }
            }
        }

        private void SaveFile(Application wordApp, Document wd)
        {
            PhraseHeader phraseHeader = dal.GetPhraseByName("Location folders");
            var firstOrDefault = phraseHeader.PhraseEntries.FirstOrDefault(pe => pe.PhraseDescription == "COA documents");
            if (firstOrDefault != null)
                _savedPath = firstOrDefault.PhraseName;
            try
            {
                string nameOfFile = "Challenge " + DateTime.Now.ToString("MM,dd,yyyy,HH,mm,ss");

                _savedPath = Path.Combine(_savedPath, nameOfFile);
                wd.SaveAs(_savedPath);
            }
            catch (Exception exception)
            {
                Process[] procs = Process.GetProcessesByName("winword");
                foreach (Process proc in procs)
                    proc.Kill();
                Common.Logger.WriteLogFile(exception);
                throw;
            }
            finally
            {
                wd.Close();
                wordApp.Application.Quit();
            }
        }

        private void PermanentFields(Document wd, Application wordApp)
        {
            Range rng = wd.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;//שדות נוספים של הפוטר
            Fields flds = rng.Fields;
            foreach (Field fld in flds)
            {
                Range rngFieldCode = fld.Code;
                String fieldText = rngFieldCode.Text;
                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    Int32 endMerge = fieldText.IndexOf("\\");
                    String fieldName = fieldText.Substring(11, endMerge - 11);
                    fieldName = fieldName.Trim();
                    if (fieldName == "replace")
                    {
                        fld.Select();
                        wordApp.Selection.TypeText(" ");
                    }
                    if (fieldName == "nuberBefor")
                    {
                        fld.Select();
                        wordApp.Selection.TypeText(" ");
                    }
                    if (fieldName == "number")
                    {
                        fld.Select();
                        string number = "";
                        if (!string.IsNullOrEmpty(_coaName))
                        {
                            number = _coaName.Remove(_coaName.IndexOf("("));
                            number = number.Replace("T", "(חלקי)").Replace("F", "");
                            wordApp.Selection.TypeText(number);
                        }

                    }
                }
            }
            foreach (Field myMergeField in wd.Fields)//שדות נוספים של גוף המסמך
            {
                Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;
                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    Int32 endMerge = fieldText.IndexOf("\\");
                    String fieldName = fieldText.Substring(11, endMerge - 11);
                    fieldName = fieldName.Trim();
                    if (fieldName == "number")
                    {
                        myMergeField.Select();
                        string number = "";
                        if (!string.IsNullOrEmpty(_coaName))
                        {
                            number = _coaName.Remove(_coaName.IndexOf("("));
                            number = number.Replace("T", "(Partial)").Replace("F", "");
                            wordApp.Selection.TypeText(number);
                        }
                    }

                }
            }
            //if (fieldName == "DrName")
            //{
            //    myMergeField.Select();
            //    if (sdg.LabInfo != null) wordApp.Selection.TypeText(sdg.LabInfo.ManagerName);
            //}
            //if (fieldName == "lab")
            //{
            //    myMergeField.Select();
            //    if (sdg.LabInfo != null) wordApp.Selection.TypeText(sdg.LabInfo.LabHebrewName);
            //}


        }
        private object GetTemplatePath(string name, bool IsGMP)
        {

            var phraseHeader = dal.GetPhraseByName("Location folders");
            var firstOrDefault = phraseHeader.PhraseEntries.FirstOrDefault(pe => pe.PhraseDescription == "COA Templates");
            if (name == "USP CHALLENGE מיוחד v3")
            {
                GetTypeOfChllenge = USP_SPECIAL;
                _typeForRow = 8;
                return IsGMP ? Path.Combine(firstOrDefault.PhraseName, "ChallengeGmp.dotx")
                           : Path.Combine(firstOrDefault.PhraseName, "Challenge.dotx");
            }
            if (name == "EP CHALLENGE v3")
            {
                GetTypeOfChllenge = EP;
                _typeForRow = 6;
                return IsGMP ? Path.Combine(firstOrDefault.PhraseName, "EpGmp.dotx")
                           : Path.Combine(firstOrDefault.PhraseName, "Ep.dotx");
            }
            if (name == "USP REG CHALLENGE v3")
            {
                GetTypeOfChllenge = USP;
                _typeForRow = 7;
                return IsGMP ? Path.Combine(firstOrDefault.PhraseName, "USPGMP.dotx")
                           : Path.Combine(firstOrDefault.PhraseName, "USP.dotx");
            }
            return "";
        }

        public string SavedPath
        {
            get { return _savedPath + ".docx"; }
        }
        private static string numToExponent(string Str, string description)
        {
            double num;
            bool isNum = double.TryParse(Str, out num);
            if (!isNum || description == "NoExp" || Str.Contains("."))
                return Str;
            if (num <= 100)
                return Str;
            var firstNum = Str.Substring(0, 1);
            var b = Str.Count() - 1;
            return ExponentString(Math.Round(Convert.ToDouble(firstNum + "." + Str.Substring(1, Str.Count() - 1)), 1) + "*10^" + b);
        }
        private static string ExponentString(string s)
        {
            if (s == "" || !s.Contains("^"))
            {
                return s;
            }
            var sSplit = s.Split('^');
            const string superscriptDigits = "\u2070\u00b9\u00b2\u00b3\u2074\u2075\u2076\u2077\u2078\u2079";
            var superscript = new string(sSplit[1].Select(x => superscriptDigits[x - '0']).ToArray());
            return sSplit[0] + superscript;
        }
    }
}
