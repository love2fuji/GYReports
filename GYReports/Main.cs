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
                EnergyData monthData = GetMonthData(RegionID, EnergyItemCode, startday, endday);
                EnergyData lastMonthData = GetMonthData(RegionID, EnergyItemCode, lastYearStartday, lastYearEndday);
                EnergyData yearData = GetYearData(RegionID, EnergyItemCode, startYear, endYear);
                EnergyData lastYearData = GetYearData(RegionID, EnergyItemCode, lastStartYear, lastEndYear);

                Console.WriteLine("月度->本期：" + monthData.Time + " 值：" + monthData.Value);
                Console.WriteLine("月度->：上年同期:" + lastMonthData.Time + " 值：" + lastMonthData.Value);

                Console.WriteLine("年度->本期：" + yearData.Time + " 值：" + yearData.Value);
                Console.WriteLine("年度->：上年同期:" + lastYearData.Time + " 值：" + lastYearData.Value);

                /*区域2-用电*/
                EnergyData R2monthData = GetMonthData(RegionID2, EnergyItemCode, startday, endday);
                EnergyData R2lastMonthData = GetMonthData(RegionID2, EnergyItemCode, lastYearStartday, lastYearEndday);
                EnergyData R2yearData = GetYearData(RegionID2, EnergyItemCode, startYear, endYear);
                EnergyData R2lastYearData = GetYearData(RegionID2, EnergyItemCode, lastStartYear, lastEndYear);

                Console.WriteLine("R2月度->本期：" + R2monthData.Time + " 值：" + R2monthData.Value);
                Console.WriteLine("R2月度->：上年同期:" + R2lastMonthData.Time + " 值：" + R2lastMonthData.Value);

                Console.WriteLine("R2年度->本期：" + R2yearData.Time + " 值：" + R2yearData.Value);
                Console.WriteLine("R2年度->：上年同期:" + R2lastYearData.Time + " 值：" + R2lastYearData.Value);
                #endregion 获取区域-用电

                #region 获取区域面积和人数
                /*获取区域面积和人数*/
                int year = endTime.Year;
                int lastYear = endTime.AddYears(-1).Year;
                //区域1
                RegionAreaPeople region1 = GetAreaAndPeople(RegionID1, EnergyItemCode, year);
                RegionAreaPeople lastRegion1 = GetAreaAndPeople(RegionID1, EnergyItemCode, lastYear);
                //区域2
                RegionAreaPeople region2 = GetAreaAndPeople(RegionID2, EnergyItemCode, year);
                RegionAreaPeople lastRegion2 = GetAreaAndPeople(RegionID2, EnergyItemCode, lastYear);

                Console.WriteLine("区域1 月度->本期面积：" + region1.Area + " 人数：" + region1.People);
                Console.WriteLine("区域1 月度->：上年同期面积:" + lastRegion1.Area + " 人数：" + lastRegion1.People);

                Console.WriteLine("区域2 月度->本期面积：" + region2.Area + " 人数：" + region2.People);
                Console.WriteLine("区域2 月度->：上年同期面积:" + lastRegion2.Area + " 人数：" + lastRegion2.People);
                #endregion 获取区域面积和人数

                #region 添加表格的内容
                this.dataGridView1.Rows.Clear();
                //第1行 总人数
                int row1 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row1].Cells[0].Value = C1Row1;
                this.dataGridView1.Rows[row1].Cells[1].Value = C2Row1;
                //当月
                this.dataGridView1.Rows[row1].Cells[2].Value = region1.People > 0 ? region1.People : '-';
                this.dataGridView1.Rows[row1].Cells[3].Value = lastRegion1.People > 0 ? lastRegion1.People : '-';
                this.dataGridView1.Rows[row1].Cells[4].Value = lastRegion1.People > 0 ? ((region1.People / lastRegion1.People) - 1) * 100 : '-';
                //当年
                this.dataGridView1.Rows[row1].Cells[5].Value = region1.People > 0 ? region1.People : '-';
                this.dataGridView1.Rows[row1].Cells[6].Value = lastRegion1.People > 0 ? lastRegion1.People : '-';
                this.dataGridView1.Rows[row1].Cells[7].Value = lastRegion1.People > 0 ? ((region1.People / lastRegion1.People) - 1) * 100 : '-';
                this.dataGridView1.Rows[row1].Cells[8].Value = "";
                this.dataGridView1.Rows[row1].Cells[9].Value = "";
                this.dataGridView1.Rows[row1].Cells[10].Value = "";

                //第2行 办公大楼用电
                int row2 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row2].Cells[0].Value = C1Row2;
                this.dataGridView1.Rows[row2].Cells[1].Value = C2Row2;
                //当月用能
                this.dataGridView1.Rows[row2].Cells[2].Value = decimal.Round(monthData.Value, 2);
                this.dataGridView1.Rows[row2].Cells[3].Value = decimal.Round(lastMonthData.Value, 2);
                this.dataGridView1.Rows[row2].Cells[4].Value = lastMonthData.Value > 0 ? decimal.Round(((monthData.Value / lastMonthData.Value) - 1) * 100, 2) : 0;
                //当年用能
                this.dataGridView1.Rows[row2].Cells[5].Value = decimal.Round(yearData.Value, 2);
                this.dataGridView1.Rows[row2].Cells[6].Value = decimal.Round(lastYearData.Value, 2);
                this.dataGridView1.Rows[row2].Cells[7].Value = lastYearData.Value > 0 ? decimal.Round(((yearData.Value / lastYearData.Value) - 1) * 100, 2) : 0;
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
                this.dataGridView1.Rows[row3].Cells[4].Value = lastRegion1.Area > 0 ? decimal.Round(((region1.Area / lastRegion1.Area) - 1) * 100, 2) : 0;
                //当年用能
                this.dataGridView1.Rows[row3].Cells[5].Value = decimal.Round(region1.Area, 2);
                this.dataGridView1.Rows[row3].Cells[6].Value = decimal.Round(lastRegion1.Area, 2);
                this.dataGridView1.Rows[row3].Cells[7].Value = lastRegion1.Area > 0 ? decimal.Round(((region1.Area / lastRegion1.Area) - 1) * 100, 2) : 0;
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
                this.dataGridView1.Rows[row4].Cells[4].Value = R2lastMonthData.Value > 0 ? decimal.Round(((R2monthData.Value / R2lastMonthData.Value) - 1) * 100, 2) : 0;
                //当年用能
                this.dataGridView1.Rows[row4].Cells[5].Value = decimal.Round(R2yearData.Value, 2);
                this.dataGridView1.Rows[row4].Cells[6].Value = decimal.Round(R2lastYearData.Value, 2);
                this.dataGridView1.Rows[row4].Cells[7].Value = R2lastYearData.Value > 0 ? decimal.Round(((R2yearData.Value / R2lastYearData.Value) - 1) * 100, 2) : 0;
                this.dataGridView1.Rows[row4].Cells[8].Value = "";
                this.dataGridView1.Rows[row4].Cells[9].Value = "";
                this.dataGridView1.Rows[row4].Cells[10].Value = "";

                //第5行 其他用电建筑面积
                int row5 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row5].Cells[0].Value = C1Row3;
                this.dataGridView1.Rows[row5].Cells[1].Value = C2Row5;
                //当月用能
                this.dataGridView1.Rows[row5].Cells[2].Value = decimal.Round(region2.Area, 2);
                this.dataGridView1.Rows[row5].Cells[3].Value = decimal.Round(lastRegion2.Area, 2);
                this.dataGridView1.Rows[row5].Cells[4].Value = lastRegion2.Area > 0 ? decimal.Round(((region2.Area / lastRegion2.Area) - 1) * 100, 2) : 0;
                //当年用能
                this.dataGridView1.Rows[row5].Cells[5].Value = decimal.Round(region2.Area, 2);
                this.dataGridView1.Rows[row5].Cells[6].Value = decimal.Round(lastRegion2.Area, 2);
                this.dataGridView1.Rows[row5].Cells[7].Value = lastRegion2.Area > 0 ? decimal.Round(((region2.Area / lastRegion2.Area) - 1) * 100, 2) : 0;
                this.dataGridView1.Rows[row5].Cells[8].Value = "";
                this.dataGridView1.Rows[row5].Cells[9].Value = "";
                this.dataGridView1.Rows[row5].Cells[10].Value = "";

                //第6行 其他用电建筑面积
                int row6 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row6].Cells[0].Value = C2Row6;
                this.dataGridView1.Rows[row6].Cells[1].Value = C2Row6;
                //当月用能
                this.dataGridView1.Rows[row6].Cells[2].Value = decimal.Round(Convert.ToDecimal(dataGridView1.Rows[1].Cells[2].Value)
                    + Convert.ToDecimal(dataGridView1.Rows[3].Cells[2].Value), 2);
                this.dataGridView1.Rows[row6].Cells[3].Value = decimal.Round(Convert.ToDecimal(dataGridView1.Rows[1].Cells[3].Value)
                    + Convert.ToDecimal(dataGridView1.Rows[3].Cells[3].Value), 2);
                this.dataGridView1.Rows[row6].Cells[4].Value = Convert.ToDecimal(dataGridView1.Rows[5].Cells[3].Value) > 0 ? decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[2].Value)
                    / Convert.ToDecimal(dataGridView1.Rows[5].Cells[3].Value)) - 1) * 100, 2) : 0;
                //当年用能
                this.dataGridView1.Rows[row6].Cells[5].Value = decimal.Round(Convert.ToDecimal(dataGridView1.Rows[1].Cells[5].Value)
                    + Convert.ToDecimal(dataGridView1.Rows[3].Cells[5].Value), 2);
                this.dataGridView1.Rows[row6].Cells[6].Value = decimal.Round(Convert.ToDecimal(dataGridView1.Rows[1].Cells[6].Value)
                    + Convert.ToDecimal(dataGridView1.Rows[3].Cells[6].Value), 2);
                this.dataGridView1.Rows[row6].Cells[7].Value = Convert.ToDecimal(dataGridView1.Rows[5].Cells[6].Value) > 0 ? decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[5].Value)
                    / Convert.ToDecimal(dataGridView1.Rows[5].Cells[6].Value)) - 1) * 100, 2) : 0;
                //this.dataGridView1.Rows[row6].Cells[8].Value = "";
                //this.dataGridView1.Rows[row6].Cells[9].Value = "";
                //this.dataGridView1.Rows[row6].Cells[10].Value = "";

                //第7行 人均用电量
                int row7 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row7].Cells[0].Value = C2Row7;
                this.dataGridView1.Rows[row7].Cells[1].Value = C2Row7;
                //当月用能
                this.dataGridView1.Rows[row7].Cells[2].Value = Convert.ToDecimal(dataGridView1.Rows[0].Cells[2].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[2].Value) / Convert.ToDecimal(dataGridView1.Rows[0].Cells[2].Value))), 2) : 0;
                this.dataGridView1.Rows[row7].Cells[3].Value = Convert.ToDecimal(dataGridView1.Rows[0].Cells[3].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[3].Value) / Convert.ToDecimal(dataGridView1.Rows[0].Cells[3].Value))), 2) : 0;
                this.dataGridView1.Rows[row7].Cells[4].Value = Convert.ToDecimal(dataGridView1.Rows[6].Cells[3].Value) > 0 ? 
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[6].Cells[2].Value) / Convert.ToDecimal(dataGridView1.Rows[6].Cells[3].Value)) - 1) * 100, 2) : 0;
                //当年用能
                this.dataGridView1.Rows[row7].Cells[5].Value = Convert.ToDecimal(dataGridView1.Rows[0].Cells[5].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[5].Value) / Convert.ToDecimal(dataGridView1.Rows[0].Cells[5].Value))), 2) : 0;

                this.dataGridView1.Rows[row7].Cells[6].Value = Convert.ToDecimal(dataGridView1.Rows[0].Cells[6].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[6].Value) / Convert.ToDecimal(dataGridView1.Rows[0].Cells[6].Value))), 2) : 0;

                this.dataGridView1.Rows[row7].Cells[7].Value = Convert.ToDecimal(dataGridView1.Rows[6].Cells[6].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[6].Cells[5].Value) / Convert.ToDecimal(dataGridView1.Rows[6].Cells[6].Value)) - 1) * 100, 2) : 0;
                this.dataGridView1.Rows[row7].Cells[8].Value = "";
                this.dataGridView1.Rows[row7].Cells[9].Value = "";
                this.dataGridView1.Rows[row7].Cells[10].Value = "";

                /***第8行 建筑面积合计***/
                int row8 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row8].Cells[0].Value = C2Row8;
                this.dataGridView1.Rows[row8].Cells[1].Value = C2Row8;
                //当月用能
                this.dataGridView1.Rows[row8].Cells[2].Value = decimal.Round(Convert.ToDecimal(dataGridView1.Rows[2].Cells[2].Value)
                    + Convert.ToDecimal(dataGridView1.Rows[4].Cells[2].Value), 2);
                this.dataGridView1.Rows[row8].Cells[3].Value = decimal.Round(Convert.ToDecimal(dataGridView1.Rows[2].Cells[3].Value)
                    + Convert.ToDecimal(dataGridView1.Rows[4].Cells[3].Value), 2);
                this.dataGridView1.Rows[row8].Cells[4].Value = Convert.ToDecimal(dataGridView1.Rows[7].Cells[3].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[7].Cells[2].Value) / Convert.ToDecimal(dataGridView1.Rows[7].Cells[3].Value)) - 1) * 100, 2) : 0;
                //当年用能
                this.dataGridView1.Rows[row8].Cells[5].Value = decimal.Round(Convert.ToDecimal(dataGridView1.Rows[2].Cells[5].Value)
                    + Convert.ToDecimal(dataGridView1.Rows[4].Cells[5].Value), 2);

                this.dataGridView1.Rows[row8].Cells[6].Value = decimal.Round(Convert.ToDecimal(dataGridView1.Rows[2].Cells[6].Value)
                    + Convert.ToDecimal(dataGridView1.Rows[4].Cells[6].Value), 2);

                this.dataGridView1.Rows[row8].Cells[7].Value = Convert.ToDecimal(dataGridView1.Rows[7].Cells[6].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[7].Cells[5].Value) / Convert.ToDecimal(dataGridView1.Rows[7].Cells[6].Value)) - 1) * 100, 2) : 0;
                this.dataGridView1.Rows[row8].Cells[8].Value = "";
                this.dataGridView1.Rows[row8].Cells[9].Value = "";
                this.dataGridView1.Rows[row8].Cells[10].Value = "";

                /***第9行 单位面积用电***/
                int row9 = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[row9].Cells[0].Value = C2Row9;
                this.dataGridView1.Rows[row9].Cells[1].Value = C2Row9;
                //当月用能
                this.dataGridView1.Rows[row9].Cells[2].Value = Convert.ToDecimal(dataGridView1.Rows[7].Cells[2].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[2].Value) / Convert.ToDecimal(dataGridView1.Rows[7].Cells[2].Value))), 2) : 0;

                this.dataGridView1.Rows[row9].Cells[3].Value = Convert.ToDecimal(dataGridView1.Rows[7].Cells[3].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[3].Value) / Convert.ToDecimal(dataGridView1.Rows[7].Cells[3].Value))), 2) : 0;

                this.dataGridView1.Rows[row9].Cells[4].Value = Convert.ToDecimal(dataGridView1.Rows[8].Cells[3].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[8].Cells[2].Value) / Convert.ToDecimal(dataGridView1.Rows[8].Cells[3].Value)) - 1) * 100, 2) : 0;
                //当年用能
                this.dataGridView1.Rows[row9].Cells[5].Value = Convert.ToDecimal(dataGridView1.Rows[7].Cells[5].Value) > 0 ?
                     decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[5].Value) / Convert.ToDecimal(dataGridView1.Rows[7].Cells[5].Value))), 2) : 0;

                this.dataGridView1.Rows[row9].Cells[6].Value = Convert.ToDecimal(dataGridView1.Rows[7].Cells[6].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[5].Cells[6].Value) / Convert.ToDecimal(dataGridView1.Rows[7].Cells[6].Value))), 2) : 0;

                this.dataGridView1.Rows[row9].Cells[7].Value = Convert.ToDecimal(dataGridView1.Rows[8].Cells[6].Value) > 0 ?
                    decimal.Round(((Convert.ToDecimal(dataGridView1.Rows[8].Cells[5].Value) / Convert.ToDecimal(dataGridView1.Rows[8].Cells[6].Value)) - 1) * 100, 2) : 0;
                this.dataGridView1.Rows[row9].Cells[8].Value = "";
                this.dataGridView1.Rows[row9].Cells[9].Value = "";
                this.dataGridView1.Rows[row9].Cells[10].Value = "";


                #endregion 添加表格的内容


            }
            catch (Exception ex)
            {
                MessageBox.Show("发生异常：" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void AddOneRowData(string column1, string column2, RegionAreaPeople region, RegionAreaPeople lastRegion,
            EnergyData energyData, EnergyData lastEnergyData)
        {
            //第1行 总人数
            int row1 = this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[row1].Cells[0].Value = column1;
            this.dataGridView1.Rows[row1].Cells[1].Value = column2;
            //当月
            this.dataGridView1.Rows[row1].Cells[2].Value = region.People;
            this.dataGridView1.Rows[row1].Cells[3].Value = lastRegion.People;
            this.dataGridView1.Rows[row1].Cells[4].Value = ((region.People / lastRegion.People) - 1) * 100;
            //当年
            this.dataGridView1.Rows[row1].Cells[5].Value = region.People;
            this.dataGridView1.Rows[row1].Cells[6].Value = lastRegion.People;
            this.dataGridView1.Rows[row1].Cells[7].Value = ((region.People / lastRegion.People) - 1) * 100;
            this.dataGridView1.Rows[row1].Cells[8].Value = "";
            this.dataGridView1.Rows[row1].Cells[9].Value = "";
            this.dataGridView1.Rows[row1].Cells[10].Value = "";
        }

        /// <summary>
        /// 获取月度用能数据
        /// </summary>
        /// <param name="regionID"></param>
        /// <param name="energyItemCode"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns></returns>
        public EnergyData GetMonthData(string regionID, string energyItemCode, string startTime, string endTime)
        {
            DataTable dt = new DataTable();
            EnergyData energyData = new EnergyData();

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
            if (dt.Rows.Count > 0)
            {
                energyData.ID = dt.Rows[0]["ID"].ToString();
                energyData.Name = dt.Rows[0]["Name"].ToString();
                energyData.Time = dt.Rows[0]["Time"].ToString();
                energyData.Value = Convert.ToDecimal(dt.Rows[0]["Value"]);
            }

            return energyData;
        }
        /// <summary>
        /// 获取年度用能数据
        /// </summary>
        /// <param name="regionID"></param>
        /// <param name="energyItemCode"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns></returns>
        public EnergyData GetYearData(string regionID, string energyItemCode, string startTime, string endTime)
        {
            DataTable dt = new DataTable();
            EnergyData energyData = new EnergyData();
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

            if (dt.Rows.Count > 0)
            {
                energyData.ID = dt.Rows[0]["ID"].ToString();
                energyData.Name = dt.Rows[0]["Name"].ToString();
                energyData.Time = dt.Rows[0]["Time"].ToString();
                energyData.Value = Convert.ToDecimal(dt.Rows[0]["Value"]);
            }

            return energyData;
        }
        /// <summary>
        /// 获取区域的面积和人数
        /// </summary>
        /// <param name="regionID"></param>
        /// <param name="energyItemCode"></param>
        /// <param name="startTime"></param>
        /// <returns></returns>
        public RegionAreaPeople GetAreaAndPeople(string regionID, string energyItemCode, int startTime)
        {
            DataTable dt = new DataTable();
            RegionAreaPeople regionInfo = new RegionAreaPeople();
            string AreaPeopleSQL = @"SELECT F_ReportID AS ID, F_TotalArea AS Area, F_People AS People,F_Value AS Value
                                            FROM T_ST_ReportPlan
                                            where F_ReportID ='" + regionID + @"' 
                                            AND F_EnergyItemCode='" + energyItemCode + @"' 
                                            AND F_Year=" + startTime + @"";

            dt = SQLHelper.GetDataTable(AreaPeopleSQL);

            if (dt.Rows.Count > 0)
            {
                regionInfo.ID = dt.Rows[0]["ID"].ToString();
                regionInfo.Area = Convert.ToDecimal(dt.Rows[0]["Area"]);
                regionInfo.People = Convert.ToDecimal(dt.Rows[0]["People"]);
                regionInfo.Value = Convert.ToDecimal(dt.Rows[0]["Value"]);
            }
            return regionInfo;
        }
    }
}
