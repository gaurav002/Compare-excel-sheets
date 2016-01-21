using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel.Attributes;

namespace AttendanceTransactionCompare
{
    public class AttendanceRecords
    {
        [ExcelColumn("ChildcareAuthorization_IDNO")]
        public string ChildcareAuthorization_IDNO { get; set; }
        [ExcelColumn("AttendanceTransaction_IDNO")]
        public string AttendanceTransaction_IDNO { get; set; }
        [ExcelColumn("AttendanceSubmission_IDNO")]
        public string AttendanceSubmission_IDNO { get; set; }
        [ExcelColumn("ContractorPayment_IDNO")]
        public string ContractorPayment_IDNO { get; set; }
        [ExcelColumn("AttendanceMonth_DATE")]
        public string AttendanceMonth_DATE { get; set; }

        [ExcelColumn("TransactionType_CODE")]
        public string TransactionType_CODE { get; set; }

        [ExcelColumn("EIN_NUMB")]
        public string EIN_NUMB { get; set; }
        [ExcelColumn("SiteID_NUMB")]
        public string SiteID_NUMB { get; set; }
        [ExcelColumn("MCI_NUMB")]
        public string MCI_NUMB { get; set; }
        [ExcelColumn("FirstInserted_DTTM")]
        public string FirstInserted_DTTM { get; set; }
        [ExcelColumn("LastSaved_DTTM")]
        public string LastSaved_DTTM { get; set; }

        [ExcelColumn("AttendanceSubmission_DTTM")]
        public DateTime AttendanceSubmission_DTTM { get; set; }

        [ExcelColumn("TransactionSource")]
        public string TransactionSource { get; set; }

        public string TransacationSet { get; set; }

        [ExcelColumn("Set")]
        public int RecordSet { get; set; }
    }

    public static class Transactions
    {
        public static string AsString(this string Value)
        {
            if (!string.IsNullOrEmpty(Value) && Value != "NULL")
                return Value.Trim();
            else
                return string.Empty;
        }

        public static DateTime? AsDateTime1(this string Value)
        {
            if (!string.IsNullOrEmpty(Value) && Value != "NULL")
                return DateTime.Parse(Value);
            else
                return null;
        }
    }

    public class TransactionsCompare : IEqualityComparer<AttendanceRecords>
    {

        public bool Equals(AttendanceRecords x, AttendanceRecords y)
        {
            return x.EIN_NUMB == y.EIN_NUMB &&
                   x.SiteID_NUMB == y.SiteID_NUMB &&
                   x.MCI_NUMB == y.MCI_NUMB &&
                   x.TransactionType_CODE == y.TransactionType_CODE &&
                   x.AttendanceMonth_DATE == y.AttendanceMonth_DATE;
        }

        public int GetHashCode(AttendanceRecords obj)
        {
            return 0;
        }
    }

    public class AttendanceCompareWithEINSITEMCINUMBER : IEqualityComparer<AttendanceRecords>
    {

        public bool Equals(AttendanceRecords x, AttendanceRecords y)
        {
            return x.EIN_NUMB == y.EIN_NUMB &&
                  x.SiteID_NUMB == y.SiteID_NUMB &&
                  x.MCI_NUMB == y.MCI_NUMB;
        }

        public int GetHashCode(AttendanceRecords obj)
        {
            return 0;
        }
    }

    public class AuthorizationDistinct : IEqualityComparer<AttendanceRecords>
    {

        public bool Equals(AttendanceRecords x, AttendanceRecords y)
        {
            return x.EIN_NUMB == y.EIN_NUMB &&
                  x.SiteID_NUMB == y.SiteID_NUMB &&
                  x.MCI_NUMB == y.MCI_NUMB &&
                  x.TransactionType_CODE == y.TransactionType_CODE &&
                  x.AttendanceMonth_DATE == y.AttendanceMonth_DATE;
        }

        public int GetHashCode(AttendanceRecords obj)
        {
            return 0;
        }
    }
}
