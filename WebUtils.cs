using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace WebUtils
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class CBRApi : ICBRApi
    {
        public double GetRate(int code, DateTime date)
        {
            try
            {
                CBRDaily.DailyInfoSoapClient client = new CBRDaily.DailyInfoSoapClient();
                System.Data.DataSet ds = client.GetCursOnDate(date);
                EnumerableRowCollection<DataRow> dsEl = from num in ds.Tables[0].AsEnumerable()
                                                        where Int32.Parse(num["vCode"].ToString()) == 840
                                                        select num;
                var rv = dsEl.First()["VCurs"];
                return Convert.ToDouble(rv);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return 0.0;
            }
        }
    }

    [ComVisible(true)]
    public interface ICBRApi
    {
        double GetRate(int code, DateTime date);
    }
}
