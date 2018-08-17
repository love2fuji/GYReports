using GYReports.Commom;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GYReports
{
    public partial class Main : Form
    {
        static string SqlConnectionString = Config.GetValue("MSSQLConnect");

        DataTable table = new DataTable();
        public Main()
        {
            InitializeComponent();
        }


        private void btnBaseDataExport2Excel_Click(object sender, EventArgs e)
        {
            try
            {
                string RegionID = "000001G0010002";
                string EnergyItemCode = "01000";

                tabConServerLog.SelectedTab = tabPage1;
                //startTime = startTime.AddDays(-startTime.Day+1).AddHours(-startTime.Hour).AddMinutes(-startTime.Minute).AddSeconds(-startTime.Second);

                //月度-当月
                DateTime startTime = dtpStartTime.Value;
                string startday = startTime.AddDays(-startTime.Day + 1).ToString("yyyy-MM-dd");

                DateTime endTime = dtpStopTime.Value;
                string endday = endTime.AddMonths(1).AddDays(-endTime.Day).ToString("yyyy-MM-dd");

                //月度-去年同月
                string lastYearStartday = startTime.AddDays(-startTime.Day + 1).AddYears(-1).ToString("yyyy-MM-dd");
                string lastYearEndday = endTime.AddMonths(1).AddDays(-endTime.Day).AddYears(-1).ToString("yyyy-MM-dd");
                Console.WriteLine("月度->开始：" + startday + "\n月度->结束：" + endday);
                Console.WriteLine("月度->同期开始：" + lastYearStartday + "\n月度->同期结束：" + lastYearEndday);

                //年度-当年
                string startYear = startTime.AddDays(-startTime.Day + 1).AddMonths(-startTime.Month + 1).ToString("yyyy-MM-dd");
                string endYear = endTime.AddMonths(1).AddDays(-endTime.Day).ToString("yyyy-MM-dd");

                string lastStartYear = startTime.AddDays(-startTime.Day + 1).AddMonths(-startTime.Month + 1).AddYears(-1).ToString("yyyy-MM-dd");
                string lastEndYear = endTime.AddMonths(1).AddDays(-endTime.Day).AddYears(-1).ToString("yyyy-MM-dd");

                Console.WriteLine("年度->开始：" + startYear + "\n年度->结束：" + endYear);
                Console.WriteLine("年度->同期开始：" + lastStartYear + "\n年度->同期结束：" + lastEndYear);

                DataTable dt1 = getMonthData(RegionID, EnergyItemCode, startday, endday);
                DataTable dt2 = getMonthData(RegionID, EnergyItemCode, lastYearStartday, lastYearEndday);
                EnergyData monthData = new EnergyData();
                monthData.ID=dt1.Rows[0]["ID"].ToString();
                monthData.Name=dt1.Rows[0]["Name"].ToString();
                monthData.Time=dt1.Rows[0]["Time"].ToString();
                monthData.Value=dt1.Rows[0]["Value"].ToString();

                EnergyData lastMonthData = new EnergyData();
                lastMonthData.ID = dt2.Rows[0]["ID"].ToString();
                lastMonthData.Name = dt2.Rows[0]["Name"].ToString();
                lastMonthData.Time = dt2.Rows[0]["Time"].ToString();
                lastMonthData.Value = dt2.Rows[0]["Value"].ToString();

                Console.WriteLine("月度->本期：" + monthData.Time+ " 值："+ monthData.Value);
                Console.WriteLine("月度->：上年同期:" + lastMonthData.Time + " 值：" + lastMonthData.Value);

                //if (table.Rows.Count > 0)
                //{
                //    dgvShowRegion.DataSource = table;
                //    tabConServerLog.SelectedTab = tpgRegion;
                //}
                //else
                //{
                //    MessageBox.Show("当前查询时间段内无数据，请选择其他日期", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("查询发生异常：" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public DataTable getMonthData(string regionID, string energyItemCode, string startTime, string endTime)
        {
            DataTable dt = new DataTable();
            string MonthSQL = @"SELECT Region.F_RegionID AS ID,Region.F_RegionName AS Name 
                                    ,DATEADD(MM,DATEDIFF(MM,0,DayResult.F_StartDay),0) AS 'Time'
                                    ,SUM((CASE WHEN RegionMeter.F_Operator ='加' THEN 1 ELSE -1 END)*DayResult.F_Value * RegionMeter.F_Rate/100) AS Value
                                    FROM T_MC_MeterDayResult DayResult
                                    INNER JOIN T_ST_CircuitMeterInfo Circuit ON DayResult.F_MeterID = Circuit.F_MeterID
	                                INNER JOIN T_DT_EnergyItemDict EnergyItem ON Circuit.F_EnergyItemCode = EnergyItem.F_EnergyItemCode
	                                INNER JOIN T_ST_MeterParamInfo ParamInfo ON DayResult.F_MeterParamID = ParamInfo.F_MeterParamID
	                                INNER JOIN T_ST_RegionMeter RegionMeter ON DayResult.F_MeterID = RegionMeter.F_MeterID
	                                INNER JOIN T_ST_Region Region ON Region.F_RegionID = RegionMeter.F_RegionID
                                    WHERE Region.F_RegionID ='" + regionID + @"'
                                    AND EnergyItem.F_EnergyItemCode='" + energyItemCode + @"'
                                    AND ParamInfo.F_IsEnergyValue = 1
                                    AND DayResult.F_StartDay BETWEEN '" + startTime + @"' AND  '" + endTime + @"'
                                    GROUP BY Region.F_RegionID,Region.F_RegionName,DATEADD(MM,DATEDIFF(MM,0,DayResult.F_StartDay),0)
                                    ORDER BY ID,'Time' ASC";

            dt = SQLHelper.GetDataTable(MonthSQL);
            return dt ;
           
        }

    }
}
