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

        DataTable ShowTable = new DataTable();
        //第1列
        string C1Row1 = "员工人数";
        string C1Row2 = "大楼用电";
        string C1Row3 = "其他用电";
        //第2列
        string C2Row1 = "总人数（人）";
        string C2Row2 = "办公大楼用电(度)";
        string C2Row3 = "建筑面积(平方米)";
        string C2Row4 = "其他用电量(度)";
        string C2Row5 = "建筑面积(平方米)";
        string C2Row6 = "用电量合计（度）";
        string C2Row7 = "人均用电（度/人）";
        string C2Row8 = "建筑面积合计（平方米）";
        string C2Row9 = "单位面积用电（度/平方米）";

        public Main()
        {
            InitializeComponent();
        }


        private void btnGetReport_Click(object sender, EventArgs e)
        {
            try
            {
                string RegionID = "000001G0010002";
                string RegionID1 = "000001G0010002";
                string RegionID2 = "000001G0010005";
                string EnergyItemCode = "01000";
                //获取用户选择的时间

                //DateTime startTime = dtpStartTime.Value;
                DateTime endTime = dtpStopTime.Value;
                //tabConServerLog.SelectedTab = tabPage1;
                //startTime = startTime.AddDays(-startTime.Day+1).AddHours(-startTime.Hour).AddMinutes(-startTime.Minute).AddSeconds(-startTime.Second);

                //月度-当月
                //当月第1天
                string startday = endTime.AddDays(-endTime.Day + 1).ToString("yyyy-MM-dd");
                //当月最后一天
                //string endday = endTime.AddMonths(1).AddDays(-endTime.Day).ToString("yyyy-MM-dd");

                string endday = endTime.ToString("yyyy-MM-dd");

                //月度-去年同月
                string lastYearStartday = endTime.AddDays(-endTime.Day + 1).AddYears(-1).ToString("yyyy-MM-dd");
                string lastYearEndday = endTime.AddYears(-1).ToString("yyyy-MM-dd");

                Console.WriteLine("月度->开始：" + startday + "\n月度->结束：" + endday);
                Console.WriteLine("月度->同期开始：" + lastYearStartday + "\n月度->同期结束：" + lastYearEndday);

                //年度-当年
                /*
                string startYear = startTime.AddDays(-startTime.Day + 1).AddMonths(-startTime.Month + 1).ToString("yyyy-MM-dd");
                string endYear = endTime.AddMonths(1).AddDays(-endTime.Day).ToString("yyyy-MM-dd");

                string lastStartYear = startTime.AddDays(-startTime.Day + 1).AddMonths(-startTime.Month + 1).AddYears(-1).ToString("yyyy-MM-dd");
                string lastEndYear = endTime.AddMonths(1).AddDays(-endTime.Day).AddYears(-1).ToString("yyyy-MM-dd");
                */
                string startYear = endTime.AddDays(-endTime.Day + 1).AddMonths(-endTime.Month + 1).ToString("yyyy-MM-dd");
                string endYear = endTime.ToString("yyyy-MM-dd");

                string lastStartYear = endTime.AddDays(-endTime.Day + 1).AddMonths(-endTime.Month + 1).AddYears(-1).ToString("yyyy-MM-dd");
                string lastEndYear = endTime.AddYears(-1).ToString("yyyy-MM-dd");

                Console.WriteLine("年度->开始：" + startYear + "\n年度->结束：" + endYear);
                Console.WriteLine("年度->同期开始：" + lastStartYear + "\n年度->同期结束：" + lastEndYear);
                #region 获取区域-用电
                /*区域1-用电*/
                DataTable dt1 = GetMonthData(RegionID, EnergyItemCode, startday, endday);
                DataTable dt2 = GetMonthData(RegionID, EnergyItemCode, lastYearStartday, lastYearEndday);

                DataTable dt3 = GetYearData(RegionID, EnergyItemCode, startYear, endYear);
                DataTable dt4 = GetYearData(RegionID, EnergyItemCode, lastStartYear, lastEndYear);

                EnergyData monthData = new EnergyData();
                EnergyData lastMonthData = new EnergyData();
                EnergyData yearData = new EnergyData();
                EnergyData lastYearData = new EnergyData();

                if (dt1.Rows.Count > 0)
                {
                    monthData.ID = dt1.Rows[0]["ID"].ToString();
                    monthData.Name = dt1.Rows[0]["Name"].ToString();
                    monthData.Time = dt1.Rows[0]["Time"].ToString();
                    monthData.Value = Convert.ToDecimal(dt1.Rows[0]["Value"]);
                }

                if (dt2.Rows.Count > 0)
                {
                    lastMonthData.ID = dt2.Rows[0]["ID"].ToString();
                    lastMonthData.Name = dt2.Rows[0]["Name"].ToString();
                    lastMonthData.Time = dt2.Rows[0]["Time"].ToString();
                    lastMonthData.Value = Convert.ToDecimal(dt2.Rows[0]["Value"]);
                }

                if (dt3.Rows.Count > 0)
                {
                    yearData.ID = dt3.Rows[0]["ID"].ToString();
                    yearData.Name = dt3.Rows[0]["Name"].ToString();
                    yearData.Time = dt3.Rows[0]["Time"].ToString();
                    yearData.Value = Convert.ToDecimal(dt3.Rows[0]["Value"]);
                }

                if (dt4.Rows.Count > 0)
                {
                    lastYearData.ID = dt4.Rows[0]["ID"].ToString();
                    lastYearData.Name = dt4.Rows[0]["Name"].ToString();
                    lastYearData.Time = dt4.Rows[0]["Time"].ToString();
                    lastYearData.Value = Convert.ToDecimal(dt4.Rows[0]["Value"]);
                }


                Console.WriteLine("月度->本期：" + monthData.Time + " 值：" + monthData.Value);
                Console.WriteLine("月度->：上年同期:" + lastMonthData.Time + " 值：" + lastMonthData.Value);

                Console.WriteLine("年度->本期：" + yearData.Time + " 值：" + yearData.Value);
                Console.WriteLine("年度->：上年同期:" + lastYearData.Time + " 值：" + lastYearData.Value);

                /*区域2-用电*/
                DataTable R2dt1 = GetMonthData(RegionID2, EnergyItemCode, startday, endday);
                DataTable R2dt2 = GetMonthData(RegionID2, EnergyItemCode, lastYearStartday, lastYearEndday);

                DataTable R2dt3 = GetYearData(RegionID2, EnergyItemCode, startYear, endYear);
                DataTable R2dt4 = GetYearData(RegionID2, EnergyItemCode, lastStartYear, lastEndYear);

                EnergyData R2monthData = new EnergyData();
                EnergyData R2lastMonthData = new EnergyData();
                EnergyData R2yearData = new EnergyData();
                EnergyData R2lastYearData = new EnergyData();

                if (R2dt1.Rows.Count > 0)
                {
                    R2monthData.ID = R2dt1.Rows[0]["ID"].ToString();
                    R2monthData.Name = R2dt1.Rows[0]["Name"].ToString();
                    R2monthData.Time = R2dt1.Rows[0]["Time"].ToString();
                    R2monthData.Value = Convert.ToDecimal(R2dt1.Rows[0]["Value"]);
                }

                if (R2dt2.Rows.Count > 0)
                {
                    R2lastMonthData.ID = R2dt2.Rows[0]["ID"].ToString();
                    R2lastMonthData.Name = R2dt2.Rows[0]["Name"].ToString();
                    R2lastMonthData.Time = R2dt2.Rows[0]["Time"].ToString();
                    R2lastMonthData.Value = Convert.ToDecimal(R2dt2.Rows[0]["Value"]);
                }

                if (R2dt3.Rows.Count > 0)
                {
                    R2yearData.ID = R2dt3.Rows[0]["ID"].ToString();
                    R2yearData.Name = R2dt3.Rows[0]["Name"].ToString();
                    R2yearData.Time = R2dt3.Rows[0]["Time"].ToString();
                    R2yearData.Value = Convert.ToDecimal(R2dt3.Rows[0]["Value"]);
                }

                if (R2dt4.Rows.Count > 0)
                {
                    R2lastYearData.ID = R2dt4.Rows[0]["ID"].ToString();
                    R2lastYearData.Name = R2dt4.Rows[0]["Name"].ToString();
                    R2lastYearData.Time = R2dt4.Rows[0]["Time"].ToString();
                    R2lastYearData.Value = Convert.ToDecimal(R2dt4.Rows[0]["Value"]);
                }


                Console.WriteLine("R2月度->本期：" + R2monthData.Time + " 值：" + R2monthData.Value);
                Console.WriteLine("R2月度->：上年同期:" + R2lastMonthData.Time + " 值：" + R2lastMonthData.Value);

                Console.WriteLine("R2年度->本期：" + R2yearData.Time + " 值：" + R2yearData.Value);
                Console.WriteLine("R2年度->：上年同期:" + R2lastYearData.Time + " 值：" + R2lastYearData.Value);
                #endregion 获取区域-用电

                #region 获取区域面积和人数
                /*获取区域面积和人数*/
                int year = endTime.Year;
                int lastYear = endTime.AddYears(-1).Year;

                RegionAreaPeople region1 = new RegionAreaPeople();
                RegionAreaPeople lastRegion1 = new RegionAreaPeople();

                RegionAreaPeople region2 = new RegionAreaPeople();
                RegionAreaPeople lastRegion2 = new RegionAreaPeople();

                DataTable dt5 = GetAreaAndPeople(RegionID1, EnergyItemCode, year);
                DataTable lastDt5 = GetAreaAndPeople(RegionID1, EnergyItemCode, lastYear);

                DataTable dt6 = GetAreaAndPeople(RegionID2, EnergyItemCode, year);
                DataTable laaastDt6 = GetAreaAndPeople(RegionID2, EnergyItemCode, lastYear);
                //区域1
                if (dt5.Rows.Count > 0)
                {
                    region1.ID = dt5.Rows[0]["ID"].ToString();
                    region1.Area = Convert.ToDecimal(dt5.Rows[0]["Area"]);
                    region1.People = Convert.ToDecimal(dt5.Rows[0]["People"]);
                    region1.Value = Convert.ToDecimal(dt5.Rows[0]["Value"]);
                }

                if (lastDt5.Rows.Count > 0)
                {
                    lastRegion1.ID = lastDt5.Rows[0]["ID"].ToString();
                    lastRegion1.Area = Convert.ToDecimal(lastDt5.Rows[0]["Area"]);
                    lastRegion1.People = Convert.ToDecimal(lastDt5.Rows[0]["People"]);
                    lastRegion1.Value = Convert.ToDecimal(lastDt5.Rows[0]["Value"]);
                }
                //区域2
                if (dt6.Rows.Count > 0)
                {
                    region2.ID = dt6.Rows[0]["ID"].ToString();
                    region2.Area = Convert.ToDecimal(dt6.Rows[0]["Area"]);
                    region2.People = Convert.ToDecimal(dt6.Rows[0]["People"]);
                    region2.Value = Convert.ToDecimal(dt6.Rows[0]["Value"]);
                }

                if (laaastDt6.Rows.Count > 0)
                {
                    lastRegion2.ID = laaastDt6.Rows[0]["ID"].ToString();
                    lastRegion2.Area = Convert.ToDecimal(laaastDt6.Rows[0]["Area"]);
                    lastRegion2.People = Convert.ToDecimal(laaastDt6.Rows[0]["People"]);
                    lastRegion2.Value = Convert.ToDecimal(laaastDt6.Rows[0]["Value"]);
                }

                Console.WriteLine("区域1 月度->本期面积：" + region1.Area + " 人数：" + region1.People);
                Console.WriteLine("区域1 月度->：上年同期面积:" + lastRegion1.Area + " 人数：" + lastRegion1.People);

                Console.WriteLine("区域2 月度->本期面积：" + region2.Area + " 人数：" + region2.People);
                Console.WriteLine("区域2 月度->：上年同期面积:" + lastRegion2.Area + " 人数：" + lastRegion2.People);
                #endregion 获取区域面积和人数

                #region 添加表格的内容
                //第1行 总人数
                int row1 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row1].Cells[0].Value = C1Row1;
                this.dataGridView1.Rows[row1].Cells[1].Value = C2Row1;
                //当月
                this.dataGridView1.Rows[row1].Cells[2].Value = region1.People;
                this.dataGridView1.Rows[row1].Cells[3].Value = lastRegion1.People;
                this.dataGridView1.Rows[row1].Cells[4].Value = ((region1.People / lastRegion1.People) - 1) * 100;
                //当年
                this.dataGridView1.Rows[row1].Cells[5].Value = region1.People;
                this.dataGridView1.Rows[row1].Cells[6].Value = lastRegion1.People;
                this.dataGridView1.Rows[row1].Cells[7].Value = ((region1.People / lastRegion1.People) - 1) * 100;
                this.dataGridView1.Rows[row1].Cells[8].Value = "";
                this.dataGridView1.Rows[row1].Cells[9].Value = "";
                this.dataGridView1.Rows[row1].Cells[10].Value = "";

                //第2行 办公大楼用电
                int row2 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row2].Cells[0].Value = C1Row2;
                this.dataGridView1.Rows[row2].Cells[1].Value = C2Row2;
                //当月用能
                this.dataGridView1.Rows[row2].Cells[2].Value = decimal.Round(monthData.Value,2);
                this.dataGridView1.Rows[row2].Cells[3].Value = decimal.Round(lastMonthData.Value,2);
                this.dataGridView1.Rows[row2].Cells[4].Value = decimal.Round(((monthData.Value / lastMonthData.Value) - 1) * 100,2);
                //当年用能
                this.dataGridView1.Rows[row2].Cells[5].Value = decimal.Round(yearData.Value,2);
                this.dataGridView1.Rows[row2].Cells[6].Value = decimal.Round(lastYearData.Value,2);
                this.dataGridView1.Rows[row2].Cells[7].Value = decimal.Round(((yearData.Value / lastYearData.Value) - 1) * 100,2);
                this.dataGridView1.Rows[row2].Cells[8].Value = "";
                this.dataGridView1.Rows[row2].Cells[9].Value = "";
                this.dataGridView1.Rows[row2].Cells[10].Value = "";

                //第3行 办公大楼建筑面积
                int row3 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row3].Cells[0].Value = C1Row2;
                this.dataGridView1.Rows[row3].Cells[1].Value = C2Row3;
                //当月用能
                this.dataGridView1.Rows[row3].Cells[2].Value = decimal.Round(region1.Area, 2);
                this.dataGridView1.Rows[row3].Cells[3].Value = decimal.Round(lastRegion1.Area, 2);
                this.dataGridView1.Rows[row3].Cells[4].Value = decimal.Round(((region1.Area / lastRegion1.Area) - 1) * 100, 2);
                //当年用能
                this.dataGridView1.Rows[row3].Cells[5].Value = decimal.Round(region1.Area, 2);
                this.dataGridView1.Rows[row3].Cells[6].Value = decimal.Round(lastRegion1.Area, 2);
                this.dataGridView1.Rows[row3].Cells[7].Value = decimal.Round(((region1.Area / lastRegion1.Area) - 1) * 100, 2);
                this.dataGridView1.Rows[row3].Cells[8].Value = "";
                this.dataGridView1.Rows[row3].Cells[9].Value = "";
                this.dataGridView1.Rows[row3].Cells[10].Value = "";

                //第4行 其他用电
                int row4 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row4].Cells[0].Value = C1Row3;
                this.dataGridView1.Rows[row4].Cells[1].Value = C2Row4;
                //当月用能
                this.dataGridView1.Rows[row4].Cells[2].Value = decimal.Round(R2monthData.Value, 2);
                this.dataGridView1.Rows[row4].Cells[3].Value = decimal.Round(R2lastMonthData.Value, 2);
                this.dataGridView1.Rows[row4].Cells[4].Value = decimal.Round(((R2monthData.Value / R2lastMonthData.Value) - 1) * 100, 2);
                //当年用能
                this.dataGridView1.Rows[row4].Cells[5].Value = decimal.Round(R2yearData.Value, 2);
                this.dataGridView1.Rows[row4].Cells[6].Value = decimal.Round(R2lastYearData.Value, 2);
                this.dataGridView1.Rows[row4].Cells[7].Value = decimal.Round(((R2yearData.Value / R2lastYearData.Value) - 1) * 100, 2);
                this.dataGridView1.Rows[row4].Cells[8].Value = "";
                this.dataGridView1.Rows[row4].Cells[9].Value = "";
                this.dataGridView1.Rows[row4].Cells[10].Value = "";


                #endregion 添加表格的内容

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
                MessageBox.Show("发生异常：" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        /// <summary>
        /// 获取月度用能数据
        /// </summary>
        /// <param name="regionID"></param>
        /// <param name="energyItemCode"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns></returns>
        public DataTable GetMonthData(string regionID, string energyItemCode, string startTime, string endTime)
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
            return dt;
        }
        /// <summary>
        /// 获取年度用能数据
        /// </summary>
        /// <param name="regionID"></param>
        /// <param name="energyItemCode"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns></returns>
        public DataTable GetYearData(string regionID, string energyItemCode, string startTime, string endTime)
        {
            DataTable dt = new DataTable();
            string MonthSQL = @"SELECT Region.F_RegionID AS ID,Region.F_RegionName AS Name 
                                    ,DATEADD(YEAR,DATEDIFF(YEAR,0,DayResult.F_StartDay),0) AS 'Time'
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
                                    GROUP BY Region.F_RegionID,Region.F_RegionName,DATEADD(YEAR,DATEDIFF(YEAR,0,DayResult.F_StartDay),0)
                                    ORDER BY ID,'Time' ASC";

            dt = SQLHelper.GetDataTable(MonthSQL);
            return dt;
        }

        public DataTable GetAreaAndPeople(string regionID, string energyItemCode, int startTime)
        {
            DataTable dt = new DataTable();
            string AreaPeopleSQL = @"SELECT F_ReportID AS ID, F_TotalArea AS Area, F_People AS People,F_Value AS Value
                                            FROM T_ST_ReportPlan
                                            where F_ReportID ='" + regionID + @"' 
                                            AND F_EnergyItemCode='" + energyItemCode + @"' 
                                            AND F_Year=" + startTime + @"";

            dt = SQLHelper.GetDataTable(AreaPeopleSQL);
            return dt;
        }


    }
}
