using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Common;
using DAL;
using Microsoft.Office.Interop.Word;
using DataTable = System.Data.DataTable;
using System.Diagnostics;
using Application = Microsoft.Office.Interop.Word.Application;

namespace Generate_COA_document
{
    public class GenerateDoc
    {
        private static DataTable _dt;
        private static DataTable _dtSample;
        //ארבע השורות הראשונות בדוח
        private readonly List<string> _drAoutr = new List<string>();//שורת הסמכה הכרה
        private readonly List<string> _drHeader = new List<string>();//שורת בכותרות
        private readonly List<string> _drStandard = new List<string>();//שורת שיטה תקן
        private readonly List<string> _drspesification = new List<string>();//שורת שיטה תקן
        //29.6.15 הוספת שורת LOQ למעבדת איכות סביבה
        private readonly List<string> _dLOQ = new List<string>();//שורת LOQ
        //אותה רשימת ספסיפיקציות רלוונטית גם לאיכות סביבה 29.6.15
        //27.05.15 הילה-בגלל השינוי למעבדת מזון,יש שורות מרובות של ספסיפיקציות
        //בכל תת רשימה האיבר הראשון זה כותרת ספסיקציה
        private readonly List<List<string>> _Lstspesification = new List<List<string>>();
        //תיאור הספסיפיקציות
        private List<string> specinfo = new List<string>();
        private readonly List<string> _drUnit = new List<string>();//שורת יחידת המידה
        private readonly Dictionary<int, string> _existentHed = new Dictionary<int, string>();//מחזיק את שורת הכותרות ואת העמודה שלהם
        private bool _authoriz;
        private bool _medical;
        private int _currentCell = 2;
        private IDataLayer dal;
        private DataTable _merged2;
        private bool _newCell;
        private Sdg sdg;
        private List<Sample> Samples;
        private string _savedPath;
        private string _coaName;
        private bool _english;
        private Sample sample;
        private string p;
        public void Popa()
        {
        }
        private int group_id;

        public GenerateDoc(Sdg sdg, IDataLayer dal, string coaName, bool english)//בנאי להפקת תעודה מלאה
        {

            Init();
            this.sdg = sdg;
            group_id = Convert.ToInt32(sdg.GroupId);
            this.dal = dal;
            _coaName = coaName;
            _english = english;
            CaseSdg();

        }
        public GenerateDoc(List<Sample> samples, IDataLayer dal, string coaName, bool english)//בנאי להפקת תעודה חלקית
        {

            //   Debugger.Launch();
            Init();
            this.dal = dal;
            this.Samples = samples;
            this.sdg = samples.First().Sdg;
            group_id = Convert.ToInt32(sdg.GroupId);
            _coaName = coaName;
            _english = english;
            CaseListSample();

        }

        public GenerateDoc(Sample sample, IDataLayer dal, string coaName, bool english)
        {

            Init();
            this.sdg = sample.Sdg;
            this.dal = dal;
            _coaName = coaName;
            _english = english;
            group_id = Convert.ToInt32(sdg.GroupId);
            CaseSdg1Sample(sample);

        }

        public void CancelSimlarCoa()
        {
            int i = Convert.ToInt32(_coaName.Substring(_coaName.IndexOf("(") + 1, 1));
            if (i > 1)
            {
                string coaToCancelname = _coaName;
                var coaToCancelname2 = coaToCancelname.Replace("(" + i + ")", "(" + (i - 1) + ")");
                var coaToCancel = dal.GetCoaReportByName(coaToCancelname2);
                coaToCancel.Status = "X";
                if (dal != null)
                    dal.SaveChanges();
            }
        }

        public string SavedPath
        {
            get
            {
                return _savedPath + ".docx";
            }
        }

        private void CaseListSample()//פונקציה למקרה של דוח חלקי
        {
            string currentsamplename = "";
            SetConst(); //מכניס את הערכים הקבועים בדוח (כותרות ורוווחים)
            foreach (Sample sample in Samples)//לב הדוח     
            {//29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                //27.05.15 הילה- אם מעבדת מזון , נשלוף ספסיפקציה מהפרודוקט של הסמפל
                string currSpecification;
                Boolean isspecexsits = false;
                if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                {
                    currSpecification = sample.Product.Name.ToString();

                    foreach (var currlist in _Lstspesification)
                    {
                        if (currlist[0] == currSpecification)
                        {
                            isspecexsits = true;
                        }
                    }
                    if (!isspecexsits)
                    {
                        List<string> nspesification = new List<string>();
                        nspesification.Add(currSpecification);
                        nspesification.Add("");
                        _Lstspesification.Add(nspesification);
                    }
                }
                CaseSample(sample);
                currentsamplename = sample.Name;
            }
            foreach (string head in _drHeader)
            {
                _dt.Columns.Add(new DataColumn(head, Type.GetType("System.String")));//מוסיף לדטה טיבל את הכותרות
            }
            _dt.Rows.Add(_drStandard.ToArray());//מוסיף את השורה של שיטה.תקן
            //הוספת ערך LOQ מעבדת איכות סביבה
            if (sdg.LabInfo.LabLetter == "E")
            {
                _dt.Rows.Add(_dLOQ.ToArray());//הוספת שורת LOQ
            }

            _dt.Rows.Add(_drUnit.ToArray());//מוסיף את השורה של יח מידה
            //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
            //הילה 27.05.15 שינוי שורת ספסיפיקציות למעבדת מזון
            if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
            {
                foreach (var currlist in _Lstspesification)
                {
                    int unitlenght = _drUnit.Count();
                    int currlenght = currlist.Count();

                    if (currlenght > unitlenght)
                    {

                        currlist.RemoveRange(unitlenght, (currlenght - unitlenght));
                    }
                    _dt.Rows.Add(currlist.ToArray());
                }
            }
            else
            {
                _dt.Rows.Add(_drspesification.ToArray());
            }
            _dt.Rows.Add(_drAoutr.ToArray());//מוסיף את השורה של הסמכה הכרה


            var merged = new DataTable();
            if (sdg.Client != null)
            {

                //הילה, 22.3.15 שינוי הוספת שדות נוספים לפי מעבדה
                string allFields = "";

                //ashi
                var cd = dal.GetClientData(sdg.Client.ClientId, sdg.LabInfo.LabInfoId);
                if (cd != null)
                {
                    allFields = cd.U_COA_COLUMNS;
                }

                //switch (sdg.LabInfo.LabLetter )
                //{
                //    case "C":
                //        and update current dal
                //        allFields=  sdg.Client.DefaultCOA_column;//מוסיף שדות נוספים לפי בחירת הלקוח
                //        break;
                //    case "W":
                //        allFields = sdg.Client.U_DEFAULT_COA_COLUMN_W;
                //        break;
                //    case "F":
                //        allFields = sdg.Client.U_DEFAULT_COA_COLUMN_F;
                //        break;
                //}

                if (!string.IsNullOrEmpty(allFields))
                {
                    string[] splitAllFields = allFields.Split(';');
                    foreach (string fields in splitAllFields)
                    {
                        if (!string.IsNullOrEmpty(fields))
                        {
                            string[] splitField = fields.Split('@');
                            AddRowAndColumSplited<Sample>(splitField[0], splitField[1], _merged2, currentsamplename);

                        }
                    }
                }
            }
            merged.Merge(_dtSample);
            merged.Merge(_dt);
            _merged2.Rows.Add("");
            _merged2.Rows.Add("");
            if (sdg.LabInfo.LabLetter == "E")
            {//אם זה איכות סביבה נוסיף עוד שורה בשביל שורת loq
                _merged2.Rows.Add("");
            }
            //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
            //הילה 27.05.15 שינוי שורת ספסיפיקציות למעבדת מזון
            if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
            {
                foreach (var currlist in _Lstspesification)
                {
                    _merged2.Rows.Add("");
                }
            }
            else
            {
                _merged2.Rows.Add("");
            }

            _merged2.Rows.Add("");
            MergeTwoDt(merged, _merged2);//מחבר בין הטבלאות השדות הנוסיפים והטבלה של הדוח 
            DtToWord(merged);// מוציא הכל למסמך וורד 
        }
        public void Init()
        {
            _dt = new DataTable();
            _dtSample = new DataTable();
        }

        private void CaseSdg()//פונקציה למקרה של דוח מלא 
        {
            SetConst();
            //3.5.15 שינוי כך שהשדות נוספים מספר אצוה יופיע לפי סדר נכון
            //הגדרת רשימה של שמות הסמפלים לפי סדר
            var samplesName = new List<string>();
            foreach (Sample sample in sdg.Samples.Where(s => s.Status == "C").OrderBy(x => x.SampleId))
            {
                samplesName.Add(sample.Name);
                //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                //27.05.15 הילה- אם מעבדת מזון , נשלוף ספסיפקציה מהפרודוקט של הסמפל
                string currSpecification;
                Boolean isspecexsits = false;
                if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                {
                    currSpecification = sample.Product.Name.ToString();

                    foreach (var currlist in _Lstspesification)
                    {
                        if (currlist[0] == currSpecification)
                        {
                            isspecexsits = true;
                        }
                    }
                    if (!isspecexsits)
                    {
                        List<string> nspesification = new List<string>();
                        nspesification.Add(currSpecification);
                        nspesification.Add("");
                        _Lstspesification.Add(nspesification);
                    }
                }
                CaseSample(sample);
            }
            foreach (string head in _drHeader)
            {
                _dt.Columns.Add(new DataColumn(head, Type.GetType("System.String")));
            }
            _dt.Rows.Add(_drStandard.ToArray());
            //הוספת ערך LOQ מעבדת איכות סביבה
            if (sdg.LabInfo.LabLetter == "E")
            {
                _dt.Rows.Add(_dLOQ.ToArray());//הוספת שורת LOQ
            }

            _dt.Rows.Add(_drUnit.ToArray());
            //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
            //הילה 27.05.15 שינוי שורת ספסיפיקציות למעבדת מזון
            if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
            {
                foreach (var currlist in _Lstspesification)
                {
                    // MessageBox.Show(_Lstspesification.Count.ToString() + "--" + _drUnit.Count().ToString() + "---" + currlist.Count().ToString());
                    int unitlenght = _drUnit.Count();
                    int currlenght = currlist.Count();

                    if (currlenght > unitlenght)
                    {

                        currlist.RemoveRange(unitlenght, (currlenght - unitlenght));
                    }
                    _dt.Rows.Add(currlist.ToArray());
                }
            }
            else
            {
                _dt.Rows.Add(_drspesification.ToArray());
            }
            _dt.Rows.Add(_drAoutr.ToArray());

            var merged = new DataTable();

            string allFields = "";
            var cd = dal.GetClientData(sdg.Client.ClientId, sdg.LabInfo.LabInfoId);
            if (cd != null)
            {
                allFields = cd.U_COA_COLUMNS;
            }
            //switch (sdg.LabInfo.LabLetter)
            //{
            //    case "C":
            //        //and update current dal
            //        allFields = sdg.Client.DefaultCOA_column;//מוסיף שדות נוספים לפי בחירת הלקוח
            //        break;
            //    case "W":
            //        allFields = sdg.Client.U_DEFAULT_COA_COLUMN_W;
            //        break;
            //    case "F":
            //        allFields = sdg.Client.U_DEFAULT_COA_COLUMN_F;
            //        break;
            //}
            if (!string.IsNullOrEmpty(allFields))
            {
                string[] splitAllFields = allFields.Split(';');
                foreach (string fields in splitAllFields)
                {
                    if (!string.IsNullOrEmpty(fields))
                    {
                        string[] splitField = fields.Split('@');
                        AddRowAndColum<Sample>(splitField[0], splitField[1], _merged2, samplesName);
                    }
                }
            }
            merged.Merge(_dtSample);
            merged.Merge(_dt);
            _merged2.Rows.Add("");
            _merged2.Rows.Add("");
            if (sdg.LabInfo.LabLetter == "E")
            {//אם זה איכות סביבה נוסיף עוד שורה בשביל שורת loq
                _merged2.Rows.Add("");
            }
            //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
            //הילה 27.05.15 שינוי שורת ספסיפיקציות למעבדת מזון
            if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
            {
                foreach (var currlist in _Lstspesification)
                {
                    _merged2.Rows.Add("");
                }
            }
            else
            {
                _merged2.Rows.Add("");
            }

            _merged2.Rows.Add("");
            MergeTwoDt(merged, _merged2);
            DtToWord(merged);
        }

        private void CaseSdg1Sample(Sample sample)//פונקציה למקרה של דוח מלא 
        {
            SetConst();
            if (sample.Status == "C")
            { //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                //27.05.15 הילה- אם מעבדת מזון , נשלוף ספסיפקציה מהפרודוקט של הסמפל
                string currSpecification;
                Boolean isspecexsits = false;
                if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                {
                    currSpecification = sample.Product.Name.ToString();

                    foreach (var currlist in _Lstspesification)
                    {
                        if (currlist[0] == currSpecification)
                        {
                            isspecexsits = true;
                        }
                    }
                    if (!isspecexsits)
                    {
                        List<string> nspesification = new List<string>();
                        nspesification.Add(currSpecification);
                        nspesification.Add("");
                        _Lstspesification.Add(nspesification);
                    }
                }
                CaseSample(sample);

                foreach (string head in _drHeader)
                {
                    _dt.Columns.Add(new DataColumn(head, Type.GetType("System.String")));
                }
                _dt.Rows.Add(_drStandard.ToArray());
                //הוספת ערך LOQ מעבדת איכות סביבה
                if (sdg.LabInfo.LabLetter == "E")
                {
                    _dt.Rows.Add(_dLOQ.ToArray());//הוספת שורת LOQ
                }

                _dt.Rows.Add(_drUnit.ToArray());
                //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                //הילה 27.05.15 שינוי שורת ספסיפיקציות למעבדת מזון
                if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                {
                    foreach (var currlist in _Lstspesification)
                    {
                        int unitlenght = _drUnit.Count();
                        int currlenght = currlist.Count();

                        if (currlenght > unitlenght)
                        {

                            currlist.RemoveRange(unitlenght, (currlenght - unitlenght));
                        }
                        _dt.Rows.Add(currlist.ToArray());
                    }
                }
                else
                {
                    _dt.Rows.Add(_drspesification.ToArray());
                }
                _dt.Rows.Add(_drAoutr.ToArray());

                var merged = new DataTable();
                string allFields = "";
                //ashi
                var cd = dal.GetClientData(sdg.Client.ClientId, sdg.LabInfo.LabInfoId);
                if (cd != null)
                {
                    allFields = cd.U_COA_COLUMNS;
                }
                //switch (sdg.LabInfo.LabLetter)
                //{
                //    case "C":
                //        and update current dal
                //        allFields = sdg.Client.DefaultCOA_column;//מוסיף שדות נוספים לפי בחירת הלקוח
                //        break;
                //    case "W":
                //        allFields = sdg.Client.U_DEFAULT_COA_COLUMN_W;
                //        break;
                //    case "F":
                //        allFields = sdg.Client.U_DEFAULT_COA_COLUMN_F;
                //        break;
                //}
                if (!string.IsNullOrEmpty(allFields))
                {
                    string[] splitAllFields = allFields.Split(';');
                    foreach (string fields in splitAllFields)
                    {
                        if (!string.IsNullOrEmpty(fields))
                        {
                            string[] splitField = fields.Split('@');
                            AddRowAndColumForSpic<Sample>(splitField[0], splitField[1], _merged2, sample.SampleId);
                        }
                    }
                }
                merged.Merge(_dtSample);
                merged.Merge(_dt);
                _merged2.Rows.Add("");
                _merged2.Rows.Add("");
                if (sdg.LabInfo.LabLetter == "E")
                {//אם זה איכות סביבה נוסיף עוד שורה בשביל שורת loq
                    _merged2.Rows.Add("");
                }
                //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                //הילה 27.05.15 שינוי שורת ספסיפיקציות למעבדת מזון
                if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                {
                    foreach (var currlist in _Lstspesification)
                    {
                        _merged2.Rows.Add("");
                    }
                }
                else
                {
                    _merged2.Rows.Add("");
                }

                _merged2.Rows.Add("");
                MergeTwoDt(merged, _merged2);
                DtToWord(merged);
            }

        }


        private void SetConst()//השמת ערכים קבועים בטבלה הפנימית
        {
            if (_english)
            {
                _drHeader.Add("Lab number");
                _drHeader.Add("Sample description");
                // _drHeader.Add(" ");

                _drStandard.Add("Standard/method");
                _drStandard.Add("");
                //29.6.15 רק למעבדת איכות סביבה להוסיף שורת LOQ
                if (sdg.LabInfo.LabLetter == "E")
                {
                    _dLOQ.Add("LOQ");
                    _dLOQ.Add("");
                }
                // _drStandard.Add("");
                //hila 20.5.15 שינוי כך ש* על ספסיפיקציה יהיה רק למעבדת מים
                if (sdg.LabInfo.LabLetter == "W")
                {
                    _drspesification.Add("Specification *");
                    _drspesification.Add("");
                    // _drspesification.Add("");
                }
                else if (sdg.LabInfo.LabLetter == "C")
                {
                    _drspesification.Add("Specification");
                    _drspesification.Add("");
                }
                _drUnit.Add("Units");
                _drUnit.Add("");
                //_drUnit.Add("");

                _drAoutr.Add("Accreditation/recognition");
                _drAoutr.Add("");
            }
            else
            {
                _drHeader.Add("מספר מעבדה");
                _drHeader.Add("תיאור הדוגמא");
                // _drHeader.Add(" ");

                _drStandard.Add("שיטה/תקן");
                _drStandard.Add("");
                // _drStandard.Add("");
                //29.6.15 רק למעבדת איכות סביבה להוסיף שורת LOQ
                if (sdg.LabInfo.LabLetter == "E")
                {
                    _dLOQ.Add("LOQ");
                    _dLOQ.Add("");
                }
                //hila 20.5.15 שינוי כך ש* על ספסיפיקציה יהיה רק למעבדת מים
                if (sdg.LabInfo.LabLetter == "W")
                {
                    _drspesification.Add("* ספציפיקציה");
                    _drspesification.Add("");
                    // _drspesification.Add("");
                }
                else if (sdg.LabInfo.LabLetter == "C")
                {
                    _drspesification.Add("ספציפיקציה");
                    _drspesification.Add("");
                }
                _drUnit.Add("יחידת מידה");
                _drUnit.Add("");
                //_drUnit.Add("");

                _drAoutr.Add("הסמכה/הכרה");
                _drAoutr.Add("");
            }



            _merged2 = new DataTable();
            _merged2.Columns.Add(new DataColumn("stam", Type.GetType("System.String")));
        }

        private void MergeTwoDt(DataTable one, DataTable two)//פונקציה לאיחוד שתי טבלאות מידע
        {
            int n = 1;
            //  Logger.WriteLogFile("one table rows "+ one.Rows.Count.ToString()+" tow  table rows "+ two.Rows.Count.ToString(),false);
            foreach (DataColumn col in two.Columns)
            {
                if (n > 1)
                {
                    one.Columns.Add(col.ColumnName);
                    for (int i = 0; i < one.Rows.Count; i++)
                    {
                        one.Rows[i][col.ColumnName] = two.Rows[i][col.ColumnName];
                    }
                }
                n++;
            }
            //    Logger.WriteLogFile("in MergeTwoDt after first loop",false);
            for (int i = 0; i < two.Columns.Count - 1; i++)
            {
                one.Columns[(one.Columns.Count - 1)].SetOrdinal(2);
            }
        }

        private void AddRowAndColum<T>(string property, string header, DataTable dt, List<string> name)//פונקציה להוספת שדות נוספים לפי בקשת לקוח
        {

            //הפונקציה מקבלת שם של פרופרטי ומחזירה את הערך הרלוונטי
            var newCol = new DataColumn(header);

            int i = 0;
            PropertyInfo newProperty = typeof(T).GetProperty(property);
            //13.4.15 טיפול במקרה שהפרופרטי לא קיים על הישות כדי שהתעודה לא תעוף
            // בגלל השינוי של שדות נוספים מהכספת ,פיצול תאריך ושעת דיגום ל2 שדות  במקום שדה אחד שכבר לא קיים
            if (newProperty != null)
            {

                dt.Columns.Add(newCol);
                //foreach (var item in name)
                //{
                //     object obj = newProperty.GetValue(sdg.Samples.Where(s => s.Status == "C").Where(s => s.Name == item).ToList().FirstOrDefault(), null);//02/03/2014
                //    dt.Rows[j][header] = obj;
                //    i++;
                //}
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    //   Logger.WriteLogFile(property + "--" + header + "--"  +name[j]+ "--" +j.ToString(), false);


                    //3.5.15 שינוי כך שהשדות נוספים מספר אצוה יופיע לפי סדר נכון הילה
                    //object obj = newProperty.GetValue(sdg.Samples.Where(s => s.Status == "C").ToList()[i], null);//02/03/2014
                    object obj = newProperty.GetValue(sdg.Samples.Where(s => s.Status == "C").Where(s => s.Name == name[j]).ToList().FirstOrDefault(), null);//02/03/2014
                    dt.Rows[j][header] = obj;
                    i++;
                }
            }
        }
        private void AddRowAndColumSplited<T>(string property, string header, DataTable dt, string name)//פונקציה להוספת שדות נוספים לפי בקשת לקוח
        {
            //הפונקציה מקבלת שם של פרופרטי ומחזירה את הערך הרלוונטי
            var newCol = new DataColumn(header);

            int i = 0;
            PropertyInfo newProperty = typeof(T).GetProperty(property);
            //13.4.15 טיפול במקרה שהפרופרטי לא קיים על הישות כדי שהתעודה לא תעוף
            // בגלל השינוי של שדות נוספים מהכספת ,פיצול תאריך ושעת דיגום ל2 שדות  במקום שדה אחד שכבר לא קיים
            if (newProperty != null)
            {
                object obj = "";
                dt.Columns.Add(newCol);
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    obj = newProperty.GetValue(sdg.Samples.Where(s => s.Status == "C" || s.Status == "P" || s.Status == "V").Where(s => s.Name == name).ToList().FirstOrDefault(), null);//02/03/2014
                    dt.Rows[j][header] = obj;
                    i++;
                }

                string s2 = string.Format("{0},#{1}", "A", "b");
            }
        }
        private void AddRowAndColumForSpic<T>(string property, string header, DataTable dt, long sampleId)//פונקציה להוספת שדות נוספים לפי בקשת לקוח
        {
            //הפונקציה מקבלת שם של פרופרטי ומחזירה את הערך הרלוונטי
            var newCol = new DataColumn(header);

            int i = 0;
            PropertyInfo newProperty = typeof(T).GetProperty(property);
            //13.4.15 טיפול במקרה שהפרופרטי לא קיים על הישות כדי שהתעודה לא תעוף
            // בגלל השינוי של שדות נוספים מהכספת ,פיצול תאריך ושעת דיגום ל2 שדות  במקום שדה אחד שכבר לא קיים
            if (newProperty != null)
            {
                dt.Columns.Add(newCol);
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    object obj = newProperty.GetValue(dal.GetSampleByKey(sampleId), null);
                    dt.Rows[j][header] = obj;
                    i++;
                }
            }
        }
        private void CaseSample(Sample sample)
        {
            var currentSample = new List<string>(_currentCell + 1);//בונה את השורה של הסמפל שכרגע נכנס לדוח 
            string sampleName = sample.Name;
            currentSample.Add(sampleName);//מוסיף לשורת הסמפל הנוכחי את השם
            //            currentSample.Add(sample.Description);//מוסיף לשורת הסמפל הנוכחי את התאור
            currentSample.Add(sample.Description != null ? sample.Description : "");//מוסיף לשורת הסמפל הנוכחי את התאור

            currentSample.Add("");
            _merged2.Rows.Add(sampleName);

            //ספי ביקש שכל מה שנבחר יוצג במסך שיוך בדיקות חוץ מבדיקות חוזרות
            //List<Aliquot> chargeAliq = sample.Aliqouts.Where(a => a.U_CHARGE == "T").ToList();
            List<Aliquot> chargeAliq = sample.Aliqouts.Where(a => a.Retest == "F").ToList();
            var chargeAliqNotCancel = chargeAliq.Where(a => a.Status != "X").ToList();

            //מיון של סדר העמודות בתעודה לפי שדה ששמור ברמת הלקוח או המעבדה
            var orderedList = OrderCoaColumns(chargeAliqNotCancel);
            if (orderedList != null)
                chargeAliqNotCancel = orderedList;

            foreach (Aliquot aliq in chargeAliqNotCancel)
            {
                //Get test template extendeded
                //   var testTemplateEx = aliq.WorkflowNode.Workflow.TestTemplateExes.FirstOrDefault();
                var testTemplateEx = aliq.TestTemplateEx;
                if (testTemplateEx != null && !_existentHed.ContainsValue(testTemplateEx.HeadLineHebrew))
                {
                    //הוספת כותרות הבדיקה
                    _drHeader.Add(testTemplateEx.HeadLineHebrew);
                    //הוספת התקן לבדיקה
                    _drStandard.Add(testTemplateEx.Standard);
                    //הוספת ערך LOQ מעבדת איכות סביבה
                    if (sdg.LabInfo.LabLetter == "E")
                    {
                        if (testTemplateEx.U_LOQ != null)
                        {
                            _dLOQ.Add(testTemplateEx.U_LOQ);
                        }
                    }
                    //הסמכה לבדיקה
                    _drAoutr.Add(AuthorizationRecognition(testTemplateEx));
                    //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                    //הילה 27.05.15 שינוי שורת ספסיפיקציות למעבדת מזון
                    if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                    {
                        foreach (var currlist in _Lstspesification)
                        {
                            if (currlist[0] == sample.Product.Name.ToString())
                            {
                                currlist.Add(GetSpecificationValue(sample, testTemplateEx));
                            }
                        }

                    }
                    if (sdg.LabInfo.LabLetter != "F")
                    {
                        //הוספת ספסיפיקציה
                        _drspesification.Add(GetSpecificationValue(sample, testTemplateEx));
                    }
                    //בשלב זה הוא חייב להיות ריק
                    //הילה 21.4.15 שינוי כך שאולי נשים ערך ליחידת מידה כבר בשלב זה.
                    // _drUnit.Add("");//יחדת המידה נלקחת המריזלט ולכן לא נוספת פה אלה ברמה נמוכה יותר , פה מוסיפים תא ריק שעליו תיכתב התוצאה הרצויה
                    // הילה 21.4.15 במקרה שיש יחידת מידה על הסמפל ,ניקח ממנו .
                    //במידה ולא ניקח מהדיפולט אליקוט במידה ולא אז נרד עד הריזולט ריפורטד וניקח ממנו

                    if (sample.UNIT != null)
                    {
                        _drUnit.Add(ExponentString(sample.UNIT.NAME));
                    }
                    else
                    {
                        _drUnit.Add("");
                    }
                    _currentCell = _drHeader.Count;
                    _existentHed.Add(_currentCell, testTemplateEx.HeadLineHebrew);
                    _newCell = true;
                }
                else
                {
                    //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                    //הילה 27.05.15 שינוי שורת ספסיפיקציות למעבדת מזון
                    if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                    {
                        foreach (var currlist in _Lstspesification)
                        {
                            if (currlist[0] == sample.Product.Name.ToString())
                            {
                                currlist.Add(GetSpecificationValue(sample, testTemplateEx));
                            }
                        }

                    }
                    _currentCell = _existentHed.FirstOrDefault(c => c.Value == testTemplateEx.HeadLineHebrew).Key;
                    _newCell = false;
                }
                if (aliq.Parent.Count == 0)
                {
                    CaseAliq(aliq, currentSample);
                }
            }

            for (int i = _dtSample.Columns.Count; i < _drHeader.Count; i++)
            {
                _dtSample.Columns.Add(new DataColumn(_drHeader[i], Type.GetType("System.String")));
            }
            if (currentSample.Count > _drHeader.Count)
            {
                currentSample.Remove("");
            }
            _dtSample.Rows.Add(currentSample.ToArray());
        }

        private List<Aliquot> OrderCoaColumns(List<Aliquot> chargeAliqNotCancel)
        {
            var str = sdg.Client.TableAssociationOrder;
            if (string.IsNullOrEmpty(str))
            {
                str = sdg.LabInfo.TableAssociationOrder;

            }
            if (string.IsNullOrEmpty(str))
                return null;
            var split = str.Split(',');
            var ss = split.Where(x => x != "");
            long[] ia = ss.Select(n => Convert.ToInt64(n)).ToArray();


            chargeAliqNotCancel = chargeAliqNotCancel.OrderBy(x =>
            {
                return Array.IndexOf(ia, x.U_TEST_TEMPLATE_EXTENDED);
            }).ToList();
            return chargeAliqNotCancel;
        }

        private string GetSpecificationValue(Sample sample, TestTemplateEx testTemplateEx)
        {
            var entitySpecification = dal.GetSpecification("PRODUCT", sample.Product.ProductId).FirstOrDefault();

            if (entitySpecification != null)
            {

                //03.06.15 אם מעבדת מזון נוסיף את התיאור ספסיפיקציה
                //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                {

                    if (entitySpecification.Specification.DESCRIPTION != null)
                    {

                        specinfo.Add(entitySpecification.Specification.DESCRIPTION);
                    }
                }
                var specificationGrade =
                    entitySpecification.Specification.SpecificationGrade.Where(sp => sp.Grade.Name == "COA").
                        FirstOrDefault();

                if (specificationGrade != null && testTemplateEx.ResultTemplate != null)
                {
                    var specificationItem =
                        specificationGrade.SPECIFICATION_ITEM.Where(
                            si => si.ResultTemplateId == testTemplateEx.ResultTemplate.ResultTemplateId).FirstOrDefault();

                    if (specificationItem != null)
                    {
                        return ExponentString(specificationItem.Parameter);
                    }
                }
            }
            return "";
        }


        private string AuthorizationRecognition(TestTemplateEx tte)
        {
            string ret = "";
            bool iFirst = true;
            if (tte.Authorization == "T")
            {
                if (_english)
                {
                    ret = ret + " A";
                }
                else
                {
                    ret = ret + " א";
                }
                iFirst = false;
                _authoriz = true;
            }
            if (tte.IsGMP == "T")
            {
                _medical = true;
            }
            if (tte.Recognition == "T")
            {
                if (!iFirst)
                {

                    if (_english)
                    {
                        ret = ret + " B";
                    }
                    else
                    {
                        ret = ret + "  ב";
                    }
                }
                else
                {
                    if (_english)
                    {
                        ret = ret + " B";
                    }
                    else
                    {
                        ret = ret + "  ב";
                    }

                }
            }
            if (tte.IsGMP == "T")
            {
                if (!iFirst)
                {

                    if (_english)
                    {
                        ret = ret + " C";
                    }
                    else
                    {
                        ret = ret + "  ג";
                    }
                }
                else
                {
                    if (_english)
                    {
                        ret = ret + " C";
                    }
                    else
                    {
                        ret = ret + "  ג";
                    }

                }
            }
            if (ret == "")
            {
                ret = "(-)";
            }
            return ret;
        }

        private void CaseAliq(Aliquot aliq, List<string> currentSample)
        {
            //9.2.15 הילה 
            //אם הערך על האליקוט טרו אז לא נרד למטה לתוצאה
            if (aliq.U_DEFAULT_VALUE != null)
            {
                caseDefaultAliquot(aliq, currentSample);
            }
            else
            {
                foreach (Test test in aliq.Tests)
                {
                    CaseResult(test, currentSample);
                }
            }
            foreach (Aliquot child in aliq.Children)
            {
                CaseAliq(child, currentSample);
            }
        }
        //9.2.15 הילה 
        //אם הערך על האליקוט טרו אז לא נרד למטה לתוצאה
        private void caseDefaultAliquot(Aliquot aliq, List<string> currentSampleColumns)
        {
            for (int i = currentSampleColumns.Count; i < _drHeader.Count; i++)
            {
                //הוספת עמודה לשורת ה דגימה הספציפית עד לאורך שורת הכותרות
                currentSampleColumns.Add("");
            }
            string formattedunit = "";
            if (aliq.TestTemplateEx != null)
            {
                if (aliq.TestTemplateEx.UNIT != null)
                {
                    formattedunit = aliq.TestTemplateEx.UNIT.NAME;
                }
            }

            string text = aliq.U_DEFAULT_VALUE;
            if (_newCell)
            {
                currentSampleColumns[_currentCell - 1] = ExponentString(text);
                // הילה 21.4.15 במקרה שיש יחידת מידה על הסמפל ,ניקח ממנו .
                //במידה ולא ניקח מהדיפולט אליקוט במידה ולא אז נרד עד הריזולט ריפורטד וניקח ממנו
                if (_drUnit[_currentCell - 1] == "")
                {
                    _drUnit[_currentCell - 1] = ExponentString(formattedunit);
                }
            }
            else
            {
                currentSampleColumns[_currentCell - 1] = ExponentString(text);
            }
        }
        private void CaseResult(Test tests, List<string> currentSampleColumns)
        {
            bool onlyOne = true; //TODO : נשען על ההנחה שישנו רק אליקוט אחד ריפורטד
            var res = tests.Results.Where(x => x.Status != "X").OrderBy(x => x.ResultId);
            foreach (Result result in res)
            {
                if (result.REPORTED == "T" && onlyOne)
                {
                    for (int i = currentSampleColumns.Count; i < _drHeader.Count; i++)
                    {
                        //הוספת עמודה לשורת ה דגימה הספציפית עד לאורך שורת הכותרות
                        currentSampleColumns.Add("");
                    }
                    string formattedunit = "";
                    //בוטל בהמשך לשינוי מתאריך 21.4.15
                    //if (result.Test.Aliquot.Sample.UNIT != null && !string.IsNullOrEmpty(result.Test.Aliquot.Sample.UNIT.NAME))
                    //{
                    //    formattedunit = result.Test.Aliquot.Sample.UNIT.NAME;

                    //}
                    //if (result.Test.Aliquot.UNIT1 != null && !string.IsNullOrEmpty(result.Test.Aliquot.UNIT1.NAME) && formattedunit == "")
                    //{
                    //    formattedunit = result.Test.Aliquot.UNIT1.NAME;
                    //}
                    //---------------------
                    if (formattedunit == "")
                    {
                        formattedunit = result.FORMATTED_UNIT ?? "";
                    }
                    if (formattedunit == "")//הילה 3.5.15 אם אין יחידת מידה על התוצה ניקח מהטסט טמפלייט אקסטנדד
                    {
                        if (result.Test.Aliquot.TestTemplateEx != null)
                        {
                            if (result.Test.Aliquot.TestTemplateEx.UNIT != null)
                            {
                                formattedunit = result.Test.Aliquot.TestTemplateEx.UNIT.NAME;
                            }
                        }
                    }
                    string text = result.Test.Aliquot.TestTemplateEx.U_COA_PRINT_UNFORMATTED == "T" ? result.FormattedResult : numToExponent(result.FormattedResult, result.DESCRIPTION);
                    if (_newCell)
                    {
                        currentSampleColumns[_currentCell - 1] = ExponentString(text);
                        // הילה 21.4.15 במקרה שיש יחידת מידה על הסמפל ,ניקח ממנו .
                        //במידה ולא ניקח מהדיפולט אליקוט במידה ולא אז נרד עד הריזולט ריפורטד וניקח ממנו

                        if (_drUnit[_currentCell - 1] == "")
                        {
                            _drUnit[_currentCell - 1] = ExponentString(formattedunit);
                        }
                    }
                    else
                    {
                        currentSampleColumns[_currentCell - 1] = ExponentString(text);
                    }
                    onlyOne = false;
                }
            }
        }

        public void DtToWord(DataTable fds)
        {
            Object oMissing = Missing.Value;
            Object oTemplatePath = GetTemplatePath();
            var wordApp = new Application();
            var wd = new Document();
            try
            {
                wd = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            }
            catch (Exception exception)
            {
                wd.Close();
                wordApp.Application.Quit();
                // wordApp.Quit();
                // CreateExceptionString(exception);
                Logger.WriteLogFile(exception);
                throw new Exception(exception.Message);
            }
            PermanentFields(wd, wordApp);
            Paragraph wordParagraph = wd.Paragraphs.Add(ref oMissing);
            wordParagraph.ReadingOrder = WdReadingOrder.wdReadingOrderLtr;
            //הוספת ערכים לטבלאות הקימות
            TwoFirstTable(wd);
            object missing = Type.Missing;

            Table tbl = wd.Tables[2];
            int Cell = 1;
            int cellCount = 0;
            // wd.Tables[2].AutoFitBehavior(WdAutoFitBehavior.a);   
            foreach (DataColumn col in fds.Columns)
            {
                if (cellCount > 0)
                {
                    Column cm = tbl.Columns.Add(ref missing);
                    cm.AutoFit();

                }
                cellCount++;
            }
            foreach (DataColumn col in fds.Columns)
            {
                SetHeadings(tbl.Cell(1, Cell++), col.ColumnName);
            }
            if (_english)
            {
                for (int i = 1; i < tbl.Range.Paragraphs.Count; i++)
                {
                    tbl.Range.Paragraphs[i].ReadingOrder = WdReadingOrder.wdReadingOrderLtr;
                }
            }
            else
            {
                for (int i = 1; i < tbl.Range.Paragraphs.Count; i++)
                {
                    tbl.Range.Paragraphs[i].ReadingOrder = WdReadingOrder.wdReadingOrderLtr;
                }
            }

            tbl.Borders.Enable = 1;
            if (_english)
            {
                tbl.TableDirection = WdTableDirection.wdTableDirectionLtr;
            }
            else
            {
                tbl.TableDirection = WdTableDirection.wdTableDirectionRtl;
            }

            for (int i = 0; i < fds.Rows.Count; i++)
            {
                Row newRow = tbl.Rows.Add(ref missing);

                if (i == fds.Rows.Count - 2)
                {//20.5.15 הילה, שורת ספסיפיקציה תהיה בולד רק למעבדת מים
                    if (sdg.LabInfo.LabLetter == "W")
                    {
                        newRow.Range.Font.Bold = 1;
                        // newRow.Cells[1].Range.Text = "11111111111";

                        wordApp.Selection.Range.Font.Bold = 1;
                    }
                    else
                    {
                        newRow.Range.Font.Bold = 0;
                        // newRow.Cells[1].Range.Text = "11111111111";

                        wordApp.Selection.Range.Font.Bold = 0;
                    }
                }
                else
                {
                    newRow.Range.Font.Bold = 0;
                }

                for (int j = 0; j < fds.Columns.Count; j++)
                {
                    newRow.Cells[j + 1].Range.Text = fds.Rows[i][j].ToString();
                }
            }
            //   tbl.Rows[1].Shading.BackgroundPatternColor = WdColor.wdColorGray20;//צביעת שורת הכותרת בסוג של אפור אפרורי
            for (int i = tbl.Rows.Count; i > tbl.Rows.Count - 4; i--)
            {
                // tbl.Rows[i].Shading.BackgroundPatternColor = WdColor.wdColorGray35;
                tbl.Rows[i].SetHeight(12, HeightRule: WdRowHeightRule.wdRowHeightAtLeast);
            }

            // wd.Tables[2].Columns.AutoFit();  
            Range rn = wd.Range(wd.Tables[2].Rows[wd.Tables[2].Rows.Count - 3].Range.Start,
                                wd.Tables[2].Rows[wd.Tables[2].Rows.Count].Range.Start);
            rn.Select();
            rn.ParagraphFormat.KeepWithNext = (int)Microsoft.Office.Core.MsoTriState.msoTrue;


            //PhraseHeader phraseHeader = dal.GetPhraseByName("Location folders");
            //var firstOrDefault = phraseHeader.PhraseEntries.FirstOrDefault(pe => pe.PhraseDescription == "COA documents");
            //if (firstOrDefault != null)
            //{
            string save = sdg.LabInfo.U_COA_FOLDER; //firstOrDefault.PhraseName;
            string coaName4sdg = !string.IsNullOrEmpty(sdg.U_COA_FILE) ? "_" + sdg.U_COA_FILE : "";

            string nameOfFile = sdg.Name + DateTime.Now.ToString("MM,dd,yyyy,HH,mm,ss") + coaName4sdg.MakeSafeFilename('_');//Ashi 31/5/22 add coaName4sdg.MakeSafeFilename('_')


            //tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
            //tbl.Columns.AutoFit();

            tbl.PreferredWidth = 100;
            //tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
            //tbl.Columns.AutoFit();
            // tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

            // tbl.Range.AutoFormat();


            //  tbl.Columns.AutoFit();
            _savedPath = Path.Combine(save, nameOfFile);
            //}
            try
            {
                //   wd.SaveAs(_savedPath);

                wd.SaveAs(_savedPath);


            }
            catch (Exception exception)
            {
                Process[] procs = Process.GetProcessesByName("winword");
                foreach (Process proc in procs)
                    proc.Kill();

                Logger.WriteLogFile(exception);
                throw new Exception(exception.Message);
                //CreateExceptionString(exception);
            }

            wd.Close();
            wordApp.Application.Quit();




        }
        private string GetTemplatePath()//מביא את הטמפלט המתאים לפי הגדרות
        {
            string path = "";
            //PhraseHeader phraseHeader = dal.GetPhraseByName("Location folders");
            //PhraseEntry firstOrDefault;
            //if (_english)
            //{
            //    firstOrDefault = phraseHeader.PhraseEntries.FirstOrDefault(pe => pe.PhraseDescription == "COA Templates eng");
            //}
            //else
            //{
            //    firstOrDefault = phraseHeader.PhraseEntries.FirstOrDefault(pe => pe.PhraseDescription == "COA Templates");
            //}


            //if (firstOrDefault != null)
            //{
            string save = sdg.LabInfo.U_TEMPLATE_FOLDER;
            if (_english)
            {
                save = sdg.LabInfo.U_TEMPLATE_FOLDER_ENG;
            }




            if (_authoriz && _medical)
            {
                path = Path.Combine(save, "number7.dotx");// "number7.dotx"
            }
            if (!_authoriz && !_medical)
            {
                path = Path.Combine(save, "noMedicalNoAoutorize.dotx");//"noMedicalNoAoutorize.dotx"
            }
            if (_authoriz && !_medical)
            {
                path = Path.Combine(save, "noMedical.dotx");//"noMedical.dotx"
            }
            if (!_authoriz && _medical)
            {
                path = Path.Combine(save, "noAoutorize.dotx");//"noAoutorize.dotx"
            }

            //}
            return path;
        }

        private void TwoFirstTable(Document wd)
        {
            Row row = wd.Tables[1].Rows[1];
            if (sdg.Client != null)
                row.Cells[2].Range.Text = sdg.Client.Name;

            row = wd.Tables[1].Rows[2];
            row.Cells[2].Range.Text = sdg.ContactName;

            row = wd.Tables[1].Rows[3];
            row.Cells[2].Range.Text = sdg.Address;

            row = wd.Tables[1].Rows[1];
            row.Cells[5].Range.Text = sdg.Phone;


            if (sdg.LabInfo.LabLetter == "C")//קוסמטיקה
            {
                row = wd.Tables[1].Rows[2];

                //Ashi - set First email
                string email = "";
                if (!String.IsNullOrEmpty(sdg.Emai))
                    email = sdg.Emai.Split(';')[0];
                row.Cells[5].Range.Text = email;

                row = wd.Tables[1].Rows[3];
                row.Cells[5].Range.Text = sdg.SampledByOperator.Name;
                row = wd.Tables[1].Rows[1];
              //  row.Cells[8].Range.Text = Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yyyy");

                row = wd.Tables[1].Rows[2];
    
              //  row.Cells[8].Range.Text = Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yyyy");
                //נופל בשורה הזאת כי יש רק 3 שורות ופה ניגשים לשורה שלא קיימת
               //תקלה בקוסמטיקה
                //..18-07-23
                row = wd.Tables[1].Rows[3];
                if (sdg.ExternalReference != null)
                    row.Cells[8].Range.Text = sdg.ExternalReference;
            }

            if (sdg.LabInfo.LabLetter == "M" || sdg.LabInfo.LabLetter == "A")//כימיה או אלרגניים
            {
                row = wd.Tables[1].Rows[2];

                //Ashi - set First email
                string email = "";
                if (!String.IsNullOrEmpty(sdg.Emai))
                    email = sdg.Emai.Split(';')[0];
                row.Cells[5].Range.Text = email;

                row = wd.Tables[1].Rows[3];
                row.Cells[5].Range.Text = sdg.SampledByOperator.Name;
                row = wd.Tables[1].Rows[1];
                row.Cells[8].Range.Text = Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yyyy");

                row = wd.Tables[1].Rows[2];
                row.Cells[8].Range.Text = Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yyyy");
                //נופל בשורה הזאת כי יש רק 3 שורות ופה ניגשים לשורה שלא קיימת
                row = wd.Tables[1].Rows[4];
                if (sdg.ExternalReference != null)
                    row.Cells[8].Range.Text = sdg.ExternalReference;
            }

            if (sdg.LabInfo.LabLetter == "W")//מים
            {
                row = wd.Tables[1].Rows[3];//hila 28.1.15

                //Ashi - set First email
                string email = "";
                if (!String.IsNullOrEmpty(sdg.Emai))
                    email = sdg.Emai.Split(';')[0];
                row.Cells[5].Range.Text = email;

                row = wd.Tables[1].Rows[2];
                row.Cells[5].Range.Text = sdg.SampledByOperator.Name;//hila 28.1.15
                row = wd.Tables[1].Rows[1];
                if (sdg.U_TXT_SAMPLING_TIME != null)
                    row.Cells[8].Range.Text = sdg.U_TXT_SAMPLING_TIME;//hila28.1.15//1.2.15


                row = wd.Tables[1].Rows[2];

                row.Cells[8].Range.Text =
                Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yy") + " " + sdg.U_TXT_RECIVED_TIME + " " + Convert.ToDecimal(sdg.WaterTemperature).ToString("#.#") + "°" + " C";//hila 28.1.15 //1.2.15


                row = wd.Tables[1].Rows[3];
                row.Cells[8].Range.Text = Convert.ToDateTime(sdg.U_TEST_DATE_TIME).ToString("dd/MM/yy HH:mm");//hila 28.1.15

                row = wd.Tables[1].Rows[4];
                if (sdg.U_SAMPLING_SITE != null)
                    row.Cells[2].Range.Text = sdg.U_SAMPLING_SITE;
                row = wd.Tables[1].Rows[4];
                if (sdg.U_MINISTRY_OF_HEALTH != null)
                    row.Cells[5].Range.Text = sdg.U_MINISTRY_OF_HEALTH;
                row = wd.Tables[1].Rows[4];
                if (sdg.ExternalReference != null)
                    row.Cells[8].Range.Text = sdg.ExternalReference;
            }
            //הילה הוספת שינויים למעבדת מזון 26.05.15
            if (sdg.LabInfo.LabLetter == "F")//מזון
            {
                row = wd.Tables[1].Rows[2];

                //Ashi - set First email
                string email = "";
                if (!String.IsNullOrEmpty(sdg.Emai))
                    email = sdg.Emai.Split(';')[0];
                row.Cells[5].Range.Text = email;


                row = wd.Tables[1].Rows[3];
                row.Cells[5].Range.Text = sdg.SampledByOperator.Name;

                row = wd.Tables[1].Rows[2];
                row.Cells[8].Range.Text = Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yyyy");

                row = wd.Tables[1].Rows[1];
                //  row.Cells[8].Range.Text =
                //  Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yyyy") + " " + sdg.U_TXT_RECIVED_TIME + " " + Convert.ToDecimal(sdg.Temperature).ToString("#.#") + "°" + " C";

                //הילה 08.06.15 שינוי כך שהמעלות יופיעו טוב
                string currTemp = Convert.ToDecimal(sdg.Temperature).ToString("#.#");
                string newTemperature = "";
                if (!string.IsNullOrEmpty(sdg.U_FOOD_TEMPERATURE))
                {
                    if (_english)
                    {
                        PhraseHeader phrase = dal.GetPhraseByName("COA Temp");
                        PhraseEntry entry = phrase.PhraseEntries.FirstOrDefault(x => x.PhraseDescription == sdg.U_FOOD_TEMPERATURE);
                        newTemperature = entry.PHRASE_INFO;
                    }
                    else
                    {
                        newTemperature = sdg.U_FOOD_TEMPERATURE; //string.Format("{0}{1}{2}", currTemp, "°", " C");
                    }
                }

                row.Cells[8].Range.Text =
                Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yyyy") + " " + sdg.U_TXT_RECIVED_TIME + " " + newTemperature;



                row = wd.Tables[1].Rows[3];
                if (sdg.ExternalReference != null)
                    row.Cells[8].Range.Text = sdg.ExternalReference;
            }
            //מעבדת איכות סביבה 29.6.15
            if (sdg.LabInfo.LabLetter == "E")
            {
                row = wd.Tables[1].Rows[2];

                //Ashi - set First email
                string email = "";
                if (!String.IsNullOrEmpty(sdg.Emai))
                    email = sdg.Emai.Split(';')[0];
                row.Cells[5].Range.Text = email;

                row = wd.Tables[1].Rows[3];
                row.Cells[5].Range.Text = sdg.SampledByOperator.Name;

                row = wd.Tables[1].Rows[1];//אם יש תאריך סיום דיגום לתת טווח מתאריך התחלה אחרת רק תאריך התחלה
                string digomDate = Convert.ToDateTime(sdg.U_START_SAMPLING).ToString("dd/MM/yy");
                if (sdg.U_SAMPLING_TYPE == "מורכב")//אם דיגום מורכב ויש תאריך סיום ניתן טווח
                {
                    if (sdg.U_END_SAMPLING != null)
                    {
                        digomDate = Convert.ToDateTime(sdg.U_START_SAMPLING).ToString("dd/MM/yy") + " - " + Convert.ToDateTime(sdg.U_END_SAMPLING).ToString("dd/MM/yy");
                    }
                }
                row.Cells[8].Range.Text = digomDate;

                row = wd.Tables[1].Rows[2];


                row.Cells[8].Range.Text =
                Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yy") + " " + Convert.ToDecimal(sdg.WaterTemperature).ToString("#.#") + "°" + " C";


                row = wd.Tables[1].Rows[3];
                row.Cells[8].Range.Text = Convert.ToDateTime(sdg.DeliveryDate).ToString("dd/MM/yy");

                row = wd.Tables[1].Rows[4];//רמת דחיפות
                if (sdg.U_PRIORITY != null)
                {
                    var priorityPhrase = dal.GetPhraseByName("Sdg Priority").PhraseEntries.Where(x => x.PhraseName == sdg.U_PRIORITY).FirstOrDefault();
                    row.Cells[2].Range.Text = priorityPhrase.PhraseDescription;
                }
                row = wd.Tables[1].Rows[4];//סוג דיגום
                if (sdg.U_SAMPLING_TYPE != null)
                    row.Cells[5].Range.Text = sdg.U_SAMPLING_TYPE;
                row = wd.Tables[1].Rows[4];//מס הזמנה
                if (sdg.ExternalReference != null)
                    row.Cells[8].Range.Text = sdg.ExternalReference;
            }


        }
        private string getExzamDate(Sdg sdg)
        {
            foreach (var sample in sdg.Samples)
            {
                foreach (var aliq in sample.Aliqouts)
                {
                    if (aliq.ALIQUOT_NOTE != null)
                    {
                        if (aliq.ALIQUOT_NOTE.Where(x => x.SUBJECT == "דיווח התחלת עבודה").FirstOrDefault() != null)
                        {
                            var firstOrDefault = aliq.ALIQUOT_NOTE.Where(x => x.SUBJECT == "דיווח התחלת עבודה").FirstOrDefault();
                            if (firstOrDefault != null)
                                return Convert.ToDateTime(firstOrDefault.ENTRY_DATE).ToString("dd/MM/yyyy");
                        }
                    }

                }
            }
            return Convert.ToDateTime(sdg.CREATED_ON).ToString("dd/MM/yyyy");
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

                    if (sdg.LabInfo.LabLetter == "W" && fieldName == "remark")
                    {
                        fld.Select();
                        if (sdg.U_COA_REMARKS != null)
                            wordApp.Selection.TypeText(sdg.U_COA_REMARKS);
                        else
                        {
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "sdgSpecification")//hila 28.1.15
                    {
                        fld.Select();
                        //03.06.15 הילה-בגלל השינוי למעבדת מזון,יש שורות מרובות של ספסיפיקציות
                        //בכל תת רשימה האיבר הראשון זה כותרת ספסיקציה
                        //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                        if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                        {


                            string spectodoc = " ";
                            foreach (var item in specinfo.Distinct())
                            {

                                if (item == specinfo.Last())
                                {
                                    spectodoc += item;
                                }
                                else
                                {
                                    spectodoc += item + "\n";
                                }
                            }

                            wordApp.Selection.TypeText(spectodoc);


                        }
                        else
                        {


                            if (sdg.U_SPECIFICATION != null)
                            {
                                wordApp.Selection.TypeText(sdg.U_SPECIFICATION);
                            }
                            else
                            {
                                wordApp.Selection.TypeText(" ");
                            }
                        }
                    }
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
                        if (_english)
                        {
                            //tring snumber = number.Remove(number.IndexOf("("));
                            number = _coaName.Remove(_coaName.IndexOf("("));
                            number = number.Replace("T", "(Partial)").Replace("F", "");
                            //number += " (In Progress)";
                        }
                        else
                        {
                            number = _coaName.Remove(_coaName.IndexOf("("));
                            number = number.Replace("T", "(חלקי)").Replace("F", "");
                            //number += " (בעבודה)";
                        }
                        wordApp.Selection.TypeText(number);
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
                    if (fieldName == "remark")  //Ashi 26/06/23 Remove this condition -  sdg.LabInfo.LabLetter != "C" &&
                    {
                        myMergeField.Select();
                        if (sdg.U_COA_REMARKS != null)
                            wordApp.Selection.TypeText(sdg.U_COA_REMARKS);
                        else
                        {
                            wordApp.Selection.TypeText(" ");
                        }
                    }


                    if (fieldName == "sdgSpecification")//hila 28.1.15
                    {
                        myMergeField.Select();
                        //03.06.15 הילה-בגלל השינוי למעבדת מזון,יש שורות מרובות של ספסיפיקציות
                        //בכל תת רשימה האיבר הראשון זה כותרת ספסיקציה

                        //29.6.15 שורות ספסיפיקציות למעבדת איכות סביבה כמו במזון
                        if ((sdg.LabInfo.LabLetter == "F") || (sdg.LabInfo.LabLetter == "E"))
                        {
                            string spectodoc = " ";
                            foreach (var item in specinfo.Distinct())
                            {

                                if (item == specinfo.Last())
                                {
                                    spectodoc += item;
                                }
                                else
                                {
                                    spectodoc += item + "\n";
                                }
                            }
                            wordApp.Selection.TypeText(spectodoc);
                        }
                        else
                        {


                            if (sdg.U_SPECIFICATION != null)
                            {
                                wordApp.Selection.TypeText(sdg.U_SPECIFICATION);
                            }
                            else
                            {
                                wordApp.Selection.TypeText(" ");
                            }
                        }
                    }

                    if (fieldName == "number")
                    {
                        myMergeField.Select();
                        string number = _coaName.Remove(_coaName.IndexOf("("));
                        Sample sample = dal.GetSampleByName(number);
                        if (sample != null && "VP".Contains(sample.Status))
                        {
                            if (_english)
                            {
                                number += " (In Progress)";
                            }
                            else
                            {
                                number += " (בעבודה)";
                            }
                        }

                        number = number.Replace("T", "(Partial)").Replace("F", "");

                        wordApp.Selection.TypeText(number);
                    }
                    //Ashi -Delete it ,Assign in only in authorization - 29/7/18
                    if (1 == 2)
                    {
                        if (fieldName == "DrName")
                        {
                            myMergeField.Select();
                            if (sdg.LabInfo != null)
                                wordApp.Selection.TypeText(sdg.LabInfo.ManagerName);
                        }
                        if (fieldName == "lab")
                        {
                            myMergeField.Select();
                            if (sdg.LabInfo != null)
                                wordApp.Selection.TypeText(sdg.LabInfo.LabHebrewName);
                        }
                    }
                }
            }
        }

        private static void SetHeadings(Cell tblCell, string text)
        {
            tblCell.Range.Text = text;
            tblCell.Range.Font.Bold = 1;
            tblCell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }
        public static string CreateExceptionString(Exception e)
        {
            StringBuilder sb = new StringBuilder();
            CreateExceptionString(sb, e, string.Empty);
            return sb.ToString();
        }
        private string numToExponent(string Str, string description)
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
            if (string.IsNullOrEmpty(s) || !s.Contains("^"))
            {
                return s;
            }
            var sSplit = s.Split('^');
            const string superscriptDigits = "\u2070\u00b9\u00b2\u00b3\u2074\u2075\u2076\u2077\u2078\u2079";
            var superscript = new string(sSplit[1].Select(x => superscriptDigits[x - '0']).ToArray());
            return sSplit[0] + superscript;

        }


        private static void CreateExceptionString(StringBuilder sb, Exception e, string indent)
        {
            if (indent == null)
            {
                indent = string.Empty;
            }
            else if (indent.Length > 0)
            {
                sb.AppendFormat("{0} Inner ", indent);
            }

            sb.AppendFormat("Exception Found:\n{0}Type: {1}", indent, e.GetType().FullName);
            sb.AppendFormat("\n{0}Message: {1}", indent, e.Message);
            sb.AppendFormat("\n{0}Source: {1}", indent, e.Source);
            sb.AppendFormat("\n{0}Stacktrace: {1}", indent, e.StackTrace);

            if (e.InnerException != null)
            {
                sb.Append("\n");
                CreateExceptionString(sb, e.InnerException, indent + "  ");
            }
        }
    }
}
