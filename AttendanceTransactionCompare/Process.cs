using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel.Query;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace AttendanceTransactionCompare
{
    public class ProcessExcel
    {
       
        void UpdatePSSAttendanceExcelSheets()
        {

            string strPath = _excelFilePath + _excelFileName;
            string NewFileWriting = string.Format(_excelFilePath + "PSS_{0}.xlsx", DateTime.Now.ToString("MMddyyyy"));



            string PSSSheetNameStars = "PSS-Source";
            string POCSheetNameStars = "POC-Source";


            //List<AuthorizationRecords> PSSAWWCommon = new List<AuthorizationRecords>();
            //List<AuthorizationRecords> PSSAWWDifferent = new List<AuthorizationRecords>();
            //List<AuthorizationRecords> PSSAWWMissing = new List<AuthorizationRecords>();

            List<AttendanceRecords> PSSMissingStars = new List<AttendanceRecords>();
            List<AttendanceRecords> POCMissingStars = new List<AttendanceRecords>();



            var POCCollectionStars = LinqToExcel.ExcelQueryFactory.Worksheet<AttendanceRecords>(POCSheetNameStars, strPath).ToList();
            var PSSCollectionStars = LinqToExcel.ExcelQueryFactory.Worksheet<AttendanceRecords>(PSSSheetNameStars, strPath).ToList();



            int iMissingSet = 0;

       

            foreach (var PSS in POCCollectionStars)
            {
                var Getdata = PSSCollectionStars.Where(x => x.AttendanceMonth_DATE == PSS.AttendanceMonth_DATE &&
                                                        x.MCI_NUMB == PSS.MCI_NUMB &&
                     x.AttendanceSubmission_DTTM == PSS.AttendanceSubmission_DTTM &&
                                                       x.SiteID_NUMB==PSS.SiteID_NUMB && x.EIN_NUMB==PSS.EIN_NUMB && 
                                                       x.TransactionType_CODE == PSS.TransactionType_CODE).ToList();

                if (Getdata.Count == 0)
                {
                    iMissingSet++;

                    PSSMissingStars.Add(new AttendanceRecords()
                    {
                        AttendanceMonth_DATE = PSS.AttendanceMonth_DATE,
                       AttendanceSubmission_DTTM=PSS.AttendanceSubmission_DTTM,
                       AttendanceTransaction_IDNO=PSS.AttendanceTransaction_IDNO,
                       ChildcareAuthorization_IDNO=PSS.ChildcareAuthorization_IDNO,
                       ContractorPayment_IDNO=PSS.ContractorPayment_IDNO,
                       EIN_NUMB=PSS.EIN_NUMB,
                       SiteID_NUMB=PSS.SiteID_NUMB,
                       MCI_NUMB = PSS.MCI_NUMB,
                       TransactionType_CODE = PSS.TransactionType_CODE
                    });


                }

            }
            FileInfo Files = new FileInfo(strPath);
            ExcelPackage ExcelPack = new ExcelPackage(Files);

            if (PSSMissingStars.Count > 0)
            {
                int iRow = 1;

                var PSSExcelSheet = ExcelPack.Workbook.Worksheets.Add("POC not in PSS");
                PSSExcelSheet.Cells["A" + iRow.ToString()].Value = "AttendanceTransaction_IDNO";
                PSSExcelSheet.Cells["B" + iRow.ToString()].Value = "AttendanceSubmission_DTTM";
                PSSExcelSheet.Cells["C" + iRow.ToString()].Value = "EIN_NUMB";
                PSSExcelSheet.Cells["D" + iRow.ToString()].Value = "MCI_NUMB";
                PSSExcelSheet.Cells["E" + iRow.ToString()].Value = "AttendanceMonth_DATE";
                PSSExcelSheet.Cells["F" + iRow.ToString()].Value = "TransactionType_CODE";
                PSSExcelSheet.Cells["G" + iRow.ToString()].Value = "SiteID_NUMB";
                PSSExcelSheet.Cells["H" + iRow.ToString()].Value = "ContractorPayment_IDNO";
                PSSExcelSheet.Cells["I" + iRow.ToString()].Value = "ChildCareAuthorization_IDNO";

                foreach (var obj in PSSMissingStars)
                {
                    iRow++;
                    PSSExcelSheet.Cells["A" + iRow.ToString()].Value = obj.AttendanceTransaction_IDNO;
                    PSSExcelSheet.Cells["B" + iRow.ToString()].Value = obj.AttendanceSubmission_DTTM;
                    PSSExcelSheet.Cells["C" + iRow.ToString()].Value = obj.EIN_NUMB;
                    PSSExcelSheet.Cells["D" + iRow.ToString()].Value = obj.MCI_NUMB;
                    PSSExcelSheet.Cells["E" + iRow.ToString()].Value = obj.AttendanceMonth_DATE;
                    PSSExcelSheet.Cells["F" + iRow.ToString()].Value = obj.TransactionType_CODE;
                    PSSExcelSheet.Cells["G" + iRow.ToString()].Value = obj.SiteID_NUMB;
                    PSSExcelSheet.Cells["H" + iRow.ToString()].Value = obj.ContractorPayment_IDNO;
                    PSSExcelSheet.Cells["I" + iRow.ToString()].Value = obj.ChildcareAuthorization_IDNO;
                }

            }

            ExcelPack.Save();
            ExcelPack.Dispose();
 }

        private void UpdateDuplicateTransactions()
        {
            string strPath = _excelFilePath + _excelFileName;
            string NewFileWriting = string.Format(_excelFilePath + "PSS_{0}.xlsx", DateTime.Now.ToString("MMddyyyy"));
            string POCSheetName = "POC-Source";
            string PSSSheetName = "PSS-Source";
            List<AttendanceRecords> PSSPOCDuplicates = new List<AttendanceRecords>();

            var POCCollection = LinqToExcel.ExcelQueryFactory.Worksheet<AttendanceRecords>(POCSheetName, strPath).ToList();
            var PSSCollection = LinqToExcel.ExcelQueryFactory.Worksheet<AttendanceRecords>(PSSSheetName, strPath).ToList();

            var PSSGroupCollection = PSSCollection.GroupBy(x => new { x.EIN_NUMB, x.SiteID_NUMB, x.MCI_NUMB, x.AttendanceSubmission_DTTM, x.AttendanceMonth_DATE, x.TransactionType_CODE })
                  .Select(g => new
                  {
                      KeyValues = g.Key,
                      Coll = g.Select(x => new { x.EIN_NUMB, x.SiteID_NUMB, x.MCI_NUMB, x.AttendanceMonth_DATE, x.ChildcareAuthorization_IDNO, x.AttendanceSubmission_DTTM, x.TransactionType_CODE, x.AttendanceTransaction_IDNO })
                              .ToList()

                  }).ToList();

            int iCommonSet = 0;


            foreach (var PSS in PSSGroupCollection)
            {
                var GetDataPOC = POCCollection.Where(x => x.EIN_NUMB == PSS.KeyValues.EIN_NUMB &&
                                         x.SiteID_NUMB == PSS.KeyValues.SiteID_NUMB &&
                                         x.MCI_NUMB == PSS.KeyValues.MCI_NUMB &&
                                         x.AttendanceMonth_DATE == PSS.KeyValues.AttendanceMonth_DATE &&
                                         x.AttendanceSubmission_DTTM == PSS.KeyValues.AttendanceSubmission_DTTM &&
                                         x.TransactionType_CODE == PSS.KeyValues.TransactionType_CODE)
                                            .ToList();

                var GetDataPSS = PSSCollection.Where(x => x.EIN_NUMB == PSS.KeyValues.EIN_NUMB &&
                                        x.SiteID_NUMB == PSS.KeyValues.SiteID_NUMB &&
                                        x.MCI_NUMB == PSS.KeyValues.MCI_NUMB &&
                                        x.AttendanceMonth_DATE == PSS.KeyValues.AttendanceMonth_DATE &&
                                        x.AttendanceSubmission_DTTM == PSS.KeyValues.AttendanceSubmission_DTTM &&
                                        x.TransactionType_CODE == PSS.KeyValues.TransactionType_CODE)
                                           .ToList();

                if (GetDataPSS.Count > 1 || GetDataPOC.Count > 1)
                {
                    iCommonSet++;
                    GetDataPSS.ForEach(p =>
                    {
                        PSSPOCDuplicates.Add(new AttendanceRecords()
                        {
                            ChildcareAuthorization_IDNO = p.ChildcareAuthorization_IDNO,
                            EIN_NUMB = p.EIN_NUMB,
                            SiteID_NUMB = p.SiteID_NUMB,
                            MCI_NUMB = p.MCI_NUMB,
                            AttendanceMonth_DATE = p.AttendanceMonth_DATE,
                            AttendanceSubmission_DTTM = p.AttendanceSubmission_DTTM,
                            TransactionType_CODE = p.TransactionType_CODE,
                            AttendanceTransaction_IDNO = p.AttendanceTransaction_IDNO,
                            TransactionSource = "PSS",
                            TransacationSet = iCommonSet.ToString()
                        });

                    });


                    GetDataPOC.ForEach(x =>
                    {
                        PSSPOCDuplicates.Add(new AttendanceRecords()
                        {
                            ChildcareAuthorization_IDNO = x.ChildcareAuthorization_IDNO,
                            EIN_NUMB = x.EIN_NUMB,
                            SiteID_NUMB = x.SiteID_NUMB,
                            MCI_NUMB = x.MCI_NUMB,
                            AttendanceMonth_DATE = x.AttendanceMonth_DATE,
                            AttendanceSubmission_DTTM = x.AttendanceSubmission_DTTM,
                            TransactionType_CODE = x.TransactionType_CODE,
                            AttendanceTransaction_IDNO = x.AttendanceTransaction_IDNO,
                            TransactionSource = "POC",
                            TransacationSet = iCommonSet.ToString()
                        });

                    });
                    //  }
                }
            }
            FileInfo Files = new FileInfo(strPath);
            ExcelPackage ExcelPack = new ExcelPackage(Files);

            if (PSSPOCDuplicates.Count > 0)
            {
                int iRow = 1;
                var AWWExcelSheet = ExcelPack.Workbook.Worksheets.Add("PSSPOCDuplicates");
                AWWExcelSheet.Cells["A" + iRow.ToString()].Value = "ChildCareAuthorization_IDNO";
                AWWExcelSheet.Cells["B" + iRow.ToString()].Value = "EIN_NUMB";
                AWWExcelSheet.Cells["C" + iRow.ToString()].Value = "SiteID_NUMB";
                AWWExcelSheet.Cells["D" + iRow.ToString()].Value = "MCI_NUMB";
                AWWExcelSheet.Cells["E" + iRow.ToString()].Value = "AttendanceMonth_DATE";
                AWWExcelSheet.Cells["F" + iRow.ToString()].Value = "AttendanceSubmission_DTTM";
                AWWExcelSheet.Cells["G" + iRow.ToString()].Value = "TransactionType_CODE";
                AWWExcelSheet.Cells["H" + iRow.ToString()].Value = "AttendanceTransaction_IDNO";
                AWWExcelSheet.Cells["I" + iRow.ToString()].Value = "Authorization Type";
                AWWExcelSheet.Cells["J" + iRow.ToString()].Value = "Set";
                foreach (var obj in PSSPOCDuplicates)
                {
                    iRow++;
                    AWWExcelSheet.Cells["A" + iRow.ToString()].Value = obj.ChildcareAuthorization_IDNO;
                    AWWExcelSheet.Cells["B" + iRow.ToString()].Value = obj.EIN_NUMB;
                    AWWExcelSheet.Cells["C" + iRow.ToString()].Value = obj.SiteID_NUMB;
                    AWWExcelSheet.Cells["D" + iRow.ToString()].Value = obj.MCI_NUMB;
                    AWWExcelSheet.Cells["E" + iRow.ToString()].Value = obj.AttendanceMonth_DATE;
                    AWWExcelSheet.Cells["F" + iRow.ToString()].Value = obj.AttendanceSubmission_DTTM;
                    AWWExcelSheet.Cells["G" + iRow.ToString()].Value = obj.TransactionType_CODE;
                    AWWExcelSheet.Cells["H" + iRow.ToString()].Value = obj.AttendanceTransaction_IDNO;
                    AWWExcelSheet.Cells["I" + iRow.ToString()].Value = obj.TransactionSource;
                    AWWExcelSheet.Cells["J" + iRow.ToString()].Value = obj.TransacationSet;
                }
                //ExcelPack.Save();
            }


            ExcelPack.Save();
            ExcelPack.Dispose();


        }

        void UpdatePOCAttendanceExcelSheets()
        {
            string strPath = _excelFilePath + _excelFileName;
            string NewFileWriting = string.Format(_excelFilePath + "PSS_{0}.xlsx", DateTime.Now.ToString("MMddyyyy"));



            string PSSSheetNameStars = "PSS-Source";
            string POCSheetNameStars = "POC-Source";


            List<AttendanceRecords> PSSMissingStars = new List<AttendanceRecords>();
            List<AttendanceRecords> POCMissingStars = new List<AttendanceRecords>();



            var POCCollectionStars = LinqToExcel.ExcelQueryFactory.Worksheet<AttendanceRecords>(POCSheetNameStars, strPath).ToList();
            var PSSCollectionStars = LinqToExcel.ExcelQueryFactory.Worksheet<AttendanceRecords>(PSSSheetNameStars, strPath).ToList();


           



            int iMissingSet = 0;



            foreach (var POC in PSSCollectionStars)
            {
                var Getdata = POCCollectionStars.Where(x => x.AttendanceMonth_DATE == POC.AttendanceMonth_DATE &&
                                                        x.MCI_NUMB == POC.MCI_NUMB &&
                     x.AttendanceSubmission_DTTM == POC.AttendanceSubmission_DTTM &&
                                                       x.SiteID_NUMB == POC.SiteID_NUMB && x.EIN_NUMB == POC.EIN_NUMB &&
                                                       x.TransactionType_CODE == POC.TransactionType_CODE).ToList();

                if (Getdata.Count == 0)
                {
                    iMissingSet++;

                    POCMissingStars.Add(new AttendanceRecords()
                    {
                        AttendanceMonth_DATE = POC.AttendanceMonth_DATE,
                        AttendanceSubmission_DTTM = POC.AttendanceSubmission_DTTM,
                        AttendanceTransaction_IDNO = POC.AttendanceTransaction_IDNO,
                        ChildcareAuthorization_IDNO = POC.ChildcareAuthorization_IDNO,
                        ContractorPayment_IDNO = POC.ContractorPayment_IDNO,
                        EIN_NUMB = POC.EIN_NUMB,
                        SiteID_NUMB = POC.SiteID_NUMB,
                        MCI_NUMB = POC.MCI_NUMB,
                        TransactionType_CODE = POC.TransactionType_CODE
                    });


                }

            }
            FileInfo Files = new FileInfo(strPath);
            ExcelPackage ExcelPack = new ExcelPackage(Files);

            if (POCMissingStars.Count > 0)
            {
                int iRow = 1;

                var PSSExcelSheet = ExcelPack.Workbook.Worksheets.Add("PSS not in POC");
                PSSExcelSheet.Cells["A" + iRow.ToString()].Value = "AttendanceTransaction_IDNO";
                PSSExcelSheet.Cells["B" + iRow.ToString()].Value = "AttendanceSubmission_DTTM";
                PSSExcelSheet.Cells["C" + iRow.ToString()].Value = "EIN_NUMB";
                PSSExcelSheet.Cells["D" + iRow.ToString()].Value = "MCI_NUMB";
                PSSExcelSheet.Cells["E" + iRow.ToString()].Value = "AttendanceMonth_DATE";
                PSSExcelSheet.Cells["F" + iRow.ToString()].Value = "TransactionType_CODE";
                PSSExcelSheet.Cells["G" + iRow.ToString()].Value = "SiteID_NUMB";
                PSSExcelSheet.Cells["H" + iRow.ToString()].Value = "ContractorPayment_IDNO";
                PSSExcelSheet.Cells["I" + iRow.ToString()].Value = "ChildCareAuthorization_IDNO";

                foreach (var obj in POCMissingStars)
                {
                    iRow++;
                    PSSExcelSheet.Cells["A" + iRow.ToString()].Value = obj.AttendanceTransaction_IDNO;
                    PSSExcelSheet.Cells["B" + iRow.ToString()].Value = obj.AttendanceSubmission_DTTM;
                    PSSExcelSheet.Cells["C" + iRow.ToString()].Value = obj.EIN_NUMB;
                    PSSExcelSheet.Cells["D" + iRow.ToString()].Value = obj.MCI_NUMB;
                    PSSExcelSheet.Cells["E" + iRow.ToString()].Value = obj.AttendanceMonth_DATE;
                    PSSExcelSheet.Cells["F" + iRow.ToString()].Value = obj.TransactionType_CODE;
                    PSSExcelSheet.Cells["G" + iRow.ToString()].Value = obj.SiteID_NUMB;
                    PSSExcelSheet.Cells["H" + iRow.ToString()].Value = obj.ContractorPayment_IDNO;
                    PSSExcelSheet.Cells["I" + iRow.ToString()].Value = obj.ChildcareAuthorization_IDNO;
                }

            }

            ExcelPack.Save();
            ExcelPack.Dispose();






        }

      void UpdateCommonTransactions()
     {
         string strPath = _excelFilePath + _excelFileName;
         string NewFileWriting = string.Format(_excelFilePath + "PSS_{0}.xlsx", DateTime.Now.ToString("MMddyyyy"));
         string POCSheetName = "POC-Source";
         string PSSSheetName = "PSS-Source";
         List<AttendanceRecords> PSSAWWCommon = new List<AttendanceRecords>();

         var POCCollection = LinqToExcel.ExcelQueryFactory.Worksheet<AttendanceRecords>(POCSheetName, strPath).ToList();
         var PSSCollection = LinqToExcel.ExcelQueryFactory.Worksheet<AttendanceRecords>(PSSSheetName, strPath).ToList();

         var PSSGroupCollection = PSSCollection.GroupBy(x => new { x.EIN_NUMB, x.SiteID_NUMB, x.MCI_NUMB, x.AttendanceSubmission_DTTM,x.AttendanceMonth_DATE,x.TransactionType_CODE })
               .Select(g => new
               {
                   KeyValues = g.Key,
                   Coll = g.Select(x => new { x.EIN_NUMB, x.SiteID_NUMB, x.MCI_NUMB, x.AttendanceMonth_DATE, x.ChildcareAuthorization_IDNO, x.AttendanceSubmission_DTTM,x.TransactionType_CODE,x.AttendanceTransaction_IDNO })
                           .ToList()

               }).ToList();

         int iCommonSet = 0;
       

         foreach (var PSS in PSSGroupCollection)
         {
             var GetData = POCCollection.Where(x => x.EIN_NUMB == PSS.KeyValues.EIN_NUMB &&
                                      x.SiteID_NUMB == PSS.KeyValues.SiteID_NUMB &&
                                      x.MCI_NUMB == PSS.KeyValues.MCI_NUMB &&
                                      x.AttendanceMonth_DATE == PSS.KeyValues.AttendanceMonth_DATE &&
                                      x.AttendanceSubmission_DTTM==PSS.KeyValues.AttendanceSubmission_DTTM&&
                                      x.TransactionType_CODE==PSS.KeyValues.TransactionType_CODE)
                                         .ToList();

             if(GetData.Count>0)
             { 
                 iCommonSet++;

                 PSS.Coll.ForEach(p =>
                 {
                     PSSAWWCommon.Add(new AttendanceRecords()
                     {
                         ChildcareAuthorization_IDNO = p.ChildcareAuthorization_IDNO,
                         EIN_NUMB = p.EIN_NUMB,
                         SiteID_NUMB = p.SiteID_NUMB,
                         MCI_NUMB = p.MCI_NUMB,
                         AttendanceMonth_DATE = p.AttendanceMonth_DATE,
                         AttendanceSubmission_DTTM = p.AttendanceSubmission_DTTM,
                         TransactionType_CODE=p.TransactionType_CODE,
                         AttendanceTransaction_IDNO=p.AttendanceTransaction_IDNO,
                         TransactionSource = "PSS",
                         TransacationSet = iCommonSet.ToString()
                     });

                 });


                 GetData.ForEach(x =>
                 {
                     PSSAWWCommon.Add(new AttendanceRecords()
                     {
                       ChildcareAuthorization_IDNO = x.ChildcareAuthorization_IDNO,
                         EIN_NUMB = x.EIN_NUMB,
                         SiteID_NUMB =x.SiteID_NUMB,
                         MCI_NUMB = x.MCI_NUMB,
                         AttendanceMonth_DATE = x.AttendanceMonth_DATE,
                         AttendanceSubmission_DTTM = x.AttendanceSubmission_DTTM,
                         TransactionType_CODE=x.TransactionType_CODE,
                         AttendanceTransaction_IDNO=x.AttendanceTransaction_IDNO,
                         TransactionSource = "POC",
                         TransacationSet = iCommonSet.ToString()
                     });

                 });
             }
         }

         FileInfo Files = new FileInfo(strPath);
         ExcelPackage ExcelPack = new ExcelPackage(Files);

         if (PSSAWWCommon.Count > 0)
         {
             int iRow = 1;
             var AWWExcelSheet = ExcelPack.Workbook.Worksheets.Add("PSSPOCCommon");
             AWWExcelSheet.Cells["A" + iRow.ToString()].Value = "ChildCareAuthorization_IDNO";
             AWWExcelSheet.Cells["B" + iRow.ToString()].Value = "EIN_NUMB";
             AWWExcelSheet.Cells["C" + iRow.ToString()].Value = "SiteID_NUMB";
             AWWExcelSheet.Cells["D" + iRow.ToString()].Value = "MCI_NUMB";
             AWWExcelSheet.Cells["E" + iRow.ToString()].Value = "AttendanceMonth_DATE";
             AWWExcelSheet.Cells["F" + iRow.ToString()].Value = "AttendanceSubmission_DTTM";
             AWWExcelSheet.Cells["G" + iRow.ToString()].Value = "TransactionType_CODE";
             AWWExcelSheet.Cells["H" + iRow.ToString()].Value = "AttendanceTransaction_IDNO";
             AWWExcelSheet.Cells["I" + iRow.ToString()].Value = "Authorization Type";
             AWWExcelSheet.Cells["J" + iRow.ToString()].Value = "Set";
             foreach (var obj in PSSAWWCommon)
             {
                 iRow++;
                 AWWExcelSheet.Cells["A" + iRow.ToString()].Value = obj.ChildcareAuthorization_IDNO;
                 AWWExcelSheet.Cells["B" + iRow.ToString()].Value = obj.EIN_NUMB;
                 AWWExcelSheet.Cells["C" + iRow.ToString()].Value = obj.SiteID_NUMB;
                 AWWExcelSheet.Cells["D" + iRow.ToString()].Value = obj.MCI_NUMB;
                 AWWExcelSheet.Cells["E" + iRow.ToString()].Value = obj.AttendanceMonth_DATE;
                 AWWExcelSheet.Cells["F" + iRow.ToString()].Value = obj.AttendanceSubmission_DTTM;
                 AWWExcelSheet.Cells["G" + iRow.ToString()].Value = obj.TransactionType_CODE;
                 AWWExcelSheet.Cells["H" + iRow.ToString()].Value = obj.AttendanceTransaction_IDNO;
                 AWWExcelSheet.Cells["I" + iRow.ToString()].Value = obj.TransactionSource;
                 AWWExcelSheet.Cells["J" + iRow.ToString()].Value = obj.TransacationSet;
             }
             //ExcelPack.Save();
         }


         ExcelPack.Save();
         ExcelPack.Dispose();
     }


        public void ProcessPSSData(string strPath)
        {
            _excelFileName = Path.GetFileName(strPath);
            _excelFilePath = Path.GetDirectoryName(strPath) + @"\";



            #region [Step #2: PSS Sheets]
          
            UpdatePSSAttendanceExcelSheets();
            #endregion



        }

        public void ProcessPOCData(string strPath)
        {
            _excelFileName = Path.GetFileName(strPath);
            _excelFilePath = Path.GetDirectoryName(strPath) + @"\";



            #region [Step #2: PSS Sheets]
            UpdatePOCAttendanceExcelSheets();
            #endregion



        }

        public void ProcessPOCPSSCommonData(string strPath)
        {
            _excelFileName = Path.GetFileName(strPath);
            _excelFilePath = Path.GetDirectoryName(strPath) + @"\";



            #region [Step #2: PSS Sheets]
            UpdateCommonTransactions();
            #endregion



        }

        public void ProcessPOCPSSDuplicateData(string strPath)
        {
            _excelFileName = Path.GetFileName(strPath);
            _excelFilePath = Path.GetDirectoryName(strPath) + @"\";



            #region [Step #2: PSS Sheets]
            UpdateDuplicateTransactions();
            #endregion



        }

       

        string _excelFilePath
        {
            get;
            set;
        }
        string _excelFileName
        {
            get;

            set;
        }

        //List<AttendanceRecords> _getAWWPROD
        //{
        //    get
        //    {
        //        return getSheetInfo("AWW PROD", _excelFilePath + _excelFileName).ToList();
        //    }
        //}

        //List<AttendanceRecords> _getPSSPROD
        //{
        //    get
        //    {
        //        return getSheetInfo("PSS PROD", _excelFilePath).ToList();
        //    }
        //}
    }
}
