using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using System.Collections;
using System.Configuration;




namespace 打卡异常统计
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            InitConfiguration();

            //定义一个文件打开控件
            OpenFileDialog ofd = new OpenFileDialog();
            //设置打开对话框的初始目录，默认目录为exe运行文件所在的路径
            ofd.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            //设置打开对话框的标题
            ofd.Title = "请选择需要分析的的excle文件";
            //设置打开对话框可以多选
            ofd.Multiselect = false;
            //设置对话框打开的文件类型
            ofd.Filter = "xlsx文件|*.xlsx";
            //设置文件对话框当前选定的筛选器的索引
            ofd.FilterIndex = 2;

            ofd.InitialDirectory = System.Environment.CurrentDirectory;
            //设置对话框是否记忆之前打开的目录
            ofd.RestoreDirectory = false;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if (ofd.FileName == "")
                {
                    MessageBox.Show("请选择正确的文件");
                    return;
                }

                IXLWorkbook workbook = new XLWorkbook(ofd.FileName);

                bool flag = false;
                foreach (IXLWorksheet w in workbook.Worksheets)
                {
                    if (w.Name == "蓝牙考勤")
                    {
                        flag = true;
                        break;
                    }
                }
                
                if (flag == false)
                {
                    MessageBox.Show("未找到工作表：" + Config.checkInFileName);
                    System.Windows.Forms.Application.ExitThread();
                    return;
                }

                IXLWorksheet sheet = workbook.Worksheet("蓝牙考勤");

                Employees ems = new Employees();

                foreach (IXLRow r in sheet.Rows())      //获取所有打卡记录
                {
                    bool exsitFlag = false;
                    for (int i = 0; i < 200; i++)
                    {
                        if (ems.em[i] != null)
                        {
                            if (ems.em[i].name == r.Cell(1).Value.ToString())       //找到此人的记录
                            {
                                exsitFlag = true;
                                // (ems.em[i].checkInTime.Contains(Convert.ToDateTime(r.Cell(3).Value.ToString())) == false)  //去重
                                //
                                    ems.em[i].checkInTime.Add(Convert.ToDateTime(r.Cell(3).Value.ToString()));
                                //
                                break;
                            }
                        }
                    }

                    if (exsitFlag == false)     //未找到此人的记录，新增加一个此人的
                    {
                        for (int i = 0; i < 200; i++)
                        {
                            if (ems.em[i] == null)
                            {
                                ems.em[i] = new Employee();
                                ems.em[i].name = r.Cell(1).Value.ToString();
                                ems.em[i].checkInTime.Add(Convert.ToDateTime(r.Cell(3).Value.ToString()));
                                break;
                            }
                        }
                    }
                }


                //分析异常数据
                for (int i = 0; i < 200; i++)
                {
                    if (ems.em[i] != null)
                    {
                        ems.em[i].checkInTime.Sort();
                        bool repeatNext = false;

                        for (int j = 0; j < ems.em[i].checkInTime.Count;)      //需要读取j+1
                        {
                            if (j == ems.em[i].checkInTime.Count - 1)       //只剩最后一条单一记录，不进行匹配
                            {
                                MatchingStatus status = MatchingDetectLast(ems.em[i], Convert.ToDateTime(ems.em[i].checkInTime[j]));
                                if (status != MatchingStatus.Matched)
                                {
                                    ems.em[i].CheckInException.Add(ems.em[i].checkInTime[j]);  //写入异常数据
                                    ems.em[i].CheckInExceptionComments.Add(ConvertMatchStatusToString(status));  //写入异常原因
                                }
                                break;
                            }

                            DateTime tCurrent = Convert.ToDateTime(ems.em[i].checkInTime[j]);
                            DateTime tNext = Convert.ToDateTime(ems.em[i].checkInTime[j + 1]);

                            

                            if (j > 0)  //检查前后间隔是否大于10分钟，判定重复打卡
                            {
                                DateTime tPrev = Convert.ToDateTime(ems.em[i].checkInTime[j - 1]);

                                //间隔小于10分钟为重复签到
                                TimeSpan tsCurrent = new TimeSpan(tCurrent.Ticks);
                                TimeSpan tsPrev = new TimeSpan(tPrev.Ticks);
                                TimeSpan tsNext = new TimeSpan(tNext.Ticks);
                                 
                                double differPrev = tsCurrent.Subtract(tsPrev).Duration().TotalMinutes;
                                double differNext = tsCurrent.Subtract(tsNext).Duration().TotalMinutes;

                                if (differPrev < 10)  //数据已经在上个match中使用，报错跳过
                                {
                                    if (repeatNext == false)        //如果是非repeatNext，即使上个match中的end数据，直接报错丢掉，否则继续匹配流程
                                    {
                                        ems.em[i].CheckInException.Add(ems.em[i].checkInTime[j]);  //写入异常数据
                                        ems.em[i].CheckInExceptionComments.Add(ConvertMatchStatusToString(MatchingStatus.Repeat));  //写入异常原因

                                        j += 1;
                                        continue;
                                    }
                                    else
                                    {
                                        repeatNext = false;     //重置flag，继续匹配流程
                                    }
                                }
                                else if (differNext < 10)
                                {
                                    ems.em[i].CheckInException.Add(ems.em[i].checkInTime[j]);  //写入异常数据
                                    ems.em[i].CheckInExceptionComments.Add(ConvertMatchStatusToString(MatchingStatus.Repeat));  //写入异常原因

                                    repeatNext = true;      //两个数据都作为begin都未匹配过,报错丢掉第一个，标记flag
                                    j += 1;
                                    continue;
                                }
                            }

                            MatchingStatus status1 = MatchingDetect(ems.em[i], tCurrent, tNext);

                            if (status1 == MatchingStatus.Matched)
                            {
                                j += 2;     //匹配成功跳过此pair数据
                            }
                            else
                            {
                                if (j != 0)     //非第一条记录
                                {
                                    ems.em[i].CheckInException.Add(ems.em[i].checkInTime[j]);  //写入异常数据
                                    ems.em[i].CheckInExceptionComments.Add(ConvertMatchStatusToString(status1));  //写入异常原因
                                }
                                else  //第一条数据，如果end类型则，则begin可能在上一个月的数据，跳过；否则报错
                                {
                                    //当其为beginD数据或者Unknow数据时，即使是第一条数据也报错，其他班次不好判断直接丢弃
                                    //特殊情况下，end数据也会出现Unknow,也丢弃,例如time为0:00-earlsetA之间，会返回Unknow
                                    if (status1 == MatchingStatus.UnMatchedEndD)
                                    {
                                        ems.em[i].CheckInException.Add(ems.em[i].checkInTime[j]);  //写入异常数据
                                        ems.em[i].CheckInExceptionComments.Add(ConvertMatchStatusToString(status1));  //写入异常原因
                                    }
                                }

                                j += 1;     //匹配失败只跳过t1
                            }
                        }
                    }
                }

                //更新异常数据表
                //删除原来的统计表
                foreach(IXLWorksheet w in workbook.Worksheets)
                {
                    if (w.Name == "考勤异常统计表")
                    {
                        workbook.Worksheet("考勤异常统计表").Delete();
                        break;
                    }
                }
                //重新生成统计表
                IXLWorksheet sheetException = workbook.AddWorksheet("考勤异常统计表");
                int lineIndex = 3;

                //标题
                IXLRange range = sheetException.Range("A1:L1");
                range.Merge().Style.Font.FontName = "宋体";
                range.Style.Font.FontSize = 20;
                range.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                range.Style.Font.Bold = true;
                sheetException.Cell(1, 1).Value = "考勤异常统计表";

                //列名
                
                sheetException.Column(6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                sheetException.Row(2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                sheetException.Cell(2, 1).Value = "序号";
                sheetException.Cell(2, 2).Value = "部门";
                sheetException.Cell(2, 3).Value = "姓名";
                sheetException.Cell(2, 4).Value = "身份证号码";
                sheetException.Cell(2, 5).Value = "日期";
                sheetException.Cell(2, 6).Value = "异常提示";
                sheetException.Cell(2, 7).Value = "异常原因";
                sheetException.Cell(2, 8).Value = "直属上级签名";
                sheetException.Cell(2, 9).Value = "岗位调动";
                sheetException.Cell(2, 10).Value = "离岗/到岗时间";
                sheetException.Cell(2, 11).Value = "备注(岗位名称、薪资)";
                sheetException.Cell(2, 12).Value = "项目第一责任人签字";


                //异常数据
                sheetException.Column(3).Width = 8.88;
                sheetException.Column(4).Width = 18.88;
                sheetException.Column(5).Width = 14.5;
                sheetException.Column(6).Width = 16.38;
                sheetException.Column(7).Width = 12.5;
                sheetException.Column(8).Width = 13.25;
                sheetException.Column(9).Width = 8.88;
                sheetException.Column(10).Width = 14.38;
                sheetException.Column(11).Width = 18.5;
                sheetException.Column(12).Width = 18.5;

                for (int i = 0; i < 200; i++)
                {
                    if (ems.em[i] != null)
                    {
                        if (ems.em[i].CheckInException.Count > 0)
                        {
                            for (int j = 0; j < ems.em[i].CheckInException.Count; j++)
                            {
                                if (ems.em[i].CheckInExceptionComments[j].ToString() != ConvertMatchStatusToString(MatchingStatus.Repeat))  //重复打卡不显示
                                {
                                    sheetException.Cell(lineIndex, 3).Value = ems.em[i].name;
                                    sheetException.Cell(lineIndex, 5).Value = ems.em[i].CheckInException[j];
                                    sheetException.Cell(lineIndex, 6).Value = ems.em[i].CheckInExceptionComments[j];

                                    //输出到form
                                    string nameFix = "         ".Substring(0, (4 - ems.em[i].name.Length) * 2);
                                    richTextBoxException.Text += nameFix + ems.em[i].name + "      ";
                                    richTextBoxException.Text += String.Format("{0, -24}{1, -15}\n", ems.em[i].CheckInException[j], ems.em[i].CheckInExceptionComments[j]);

                                    lineIndex++;
                                }

                            }
                        }
                    }
                }



                //表尾
                string rg = string.Format("A{0}:L{1}", sheetException.LastRowUsed().RowNumber() + 1, sheetException.LastRowUsed().RowNumber() + 1);
                sheetException.Range(rg).Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                sheetException.Cell(sheetException.LastRowUsed().RowNumber() + 1, 1).Value = "说明：调出员工说明调至某项目任某岗位，离岗时间；调入员工说明从某项目某岗位调来任某岗位，到岗时间及薪资是否调整，员工月末所在项目负责员工考勤上报。";

                rg = string.Format("A{0}:L{1}", sheetException.LastRowUsed().RowNumber() + 1, sheetException.LastRowUsed().RowNumber() + 1);
                sheetException.Range(rg).Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                sheetException.Cell(sheetException.LastRowUsed().RowNumber() + 1, 1).Value = "制表人/日期：                                                 审批人/日期：          ";

                //加边框
                rg = string.Format("A2:L{0}", sheetException.LastRowUsed().RowNumber());
                sheetException.Range(rg).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                sheetException.Range(rg).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                //合并名称
                //sheetException.Range("C4:C8").Merge();
                //IXLRow nameRow = sheetException.Row(3);
                for (int rowIndex = 2; rowIndex < sheetException.LastRowUsed().RowNumber() - 2; )
                {
                    int start = rowIndex;
                    int end;

                    for (int index = rowIndex + 1; index <= sheetException.LastRowUsed().RowNumber() - 2; index++)
                    {
                        if (sheetException.Cell(index, 3).Value != sheetException.Cell(index - 1, 3).Value)
                        {
                            end = index - 1;
                            if (start != end)
                            {
                                string rangeIndex = string.Format("A{0}:A{1}", start, end);
                                string rangeDepartment = string.Format("B{0}:B{1}", start, end);
                                string rangeName = string.Format("C{0}:C{1}", start, end);
                                string rangeId = string.Format("D{0}:D{1}", start, end);

                                sheetException.Range(rangeIndex).Merge().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                                sheetException.Range(rangeDepartment).Merge().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                                sheetException.Range(rangeName).Merge().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                                sheetException.Range(rangeId).Merge().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                            }

                            rowIndex += end - start + 1;
                            break;
                        }

                        if (index == sheetException.LastRowUsed().RowNumber() - 2)      //注意临界值
                        {
                            rowIndex += 1;
                        }
                    }
                }

                if (checkBoxSave.Checked)
                {
                    workbook.Save();
                    //MessageBox.Show("异常表已生成！");
                }

                return;
            }

        }

        public string ConvertMatchStatusToString(MatchingStatus status)
        {
            switch(status)
            {
                case MatchingStatus.Matched: return "";
                case MatchingStatus.UnMatchedBeginA: return "早班上班未打卡";
                case MatchingStatus.UnMatchedEndA: return "早班下班未打卡";
                case MatchingStatus.UnMatchedBeginB: return "中班上班未打卡";
                case MatchingStatus.UnMatchedEndB: return "中班下班未打卡";
                case MatchingStatus.UnMatchedBeginC: return "晚班上班未打卡";
                case MatchingStatus.UnMatchedEndC: return "晚班下班未打卡";
                case MatchingStatus.UnMatchedBeginD: return "行政班上班未打卡";
                case MatchingStatus.UnMatchedEndD: return "行政班下班未打卡";
                case MatchingStatus.Unknow: return "非班次时间打卡";
                case MatchingStatus.Repeat: return "重复打卡";
                default: return "未知数据";
            }
        }

        public int ParseHour(string workingTimeString)
        {
            string tmp = workingTimeString.Substring(0, workingTimeString.IndexOf(':'));
            return Convert.ToInt32(tmp);
        }

        public int ParseMin(string workingTimeString)
        {
            string tmp = workingTimeString.Substring(workingTimeString.IndexOf(':') + 1, workingTimeString.Length - workingTimeString.IndexOf(':') - 1);
            return Convert.ToInt32(tmp);
        }

        public void GetNearlyScheduling(ProbableScheduling p, CurrentWorkingHours e, DateTime time)
        {
            DateTime tEariesst, tLatest;

            //     A班上班范围
            tEariesst = e.earliestBeginA.AddHours(-(ParseHour(Config.beginOffset)));
            tEariesst = tEariesst.AddMinutes(-(ParseMin(Config.beginOffset)));

            tLatest = e.latestBeginA.AddHours(ParseHour(Config.beginOffset));
            tLatest = tLatest.AddMinutes(ParseMin(Config.beginOffset));

            if (time > tEariesst && time <= tLatest)
            {
                p.startA = true;
                p.activeCount += 1;
            }

            //     A班下班范围
            tEariesst = e.earliestEndA.AddHours(-(ParseHour(Config.endOffset)));
            tEariesst = tEariesst.AddMinutes(-(ParseMin(Config.endOffset)));

            tLatest = e.latestEndA.AddHours(ParseHour(Config.endOffset));
            tLatest = tLatest.AddMinutes(ParseMin(Config.endOffset));

            if (time > tEariesst && time <= tLatest)
            {
                p.endA = true;
                p.activeCount += 1;
            }

            //     B班上班啊范围
            tEariesst = e.earliestBeginB.AddHours(-(ParseHour(Config.beginOffset)));
            tEariesst = tEariesst.AddMinutes(-(ParseMin(Config.beginOffset)));

            tLatest = e.latestBeginB.AddHours(ParseHour(Config.beginOffset));
            tLatest = tLatest.AddMinutes(ParseMin(Config.beginOffset));

            if (time > tEariesst && time <= tLatest)
            {
                p.startB = true;
                p.activeCount += 1;
            }

            //     B班下班范围
            tEariesst = e.earliestEndB.AddHours(-(ParseHour(Config.endOffset)));
            tEariesst = tEariesst.AddMinutes(-(ParseMin(Config.endOffset)));

            tLatest = e.latestEndB.AddHours(ParseHour(Config.endOffset));
            tLatest = tLatest.AddMinutes(ParseMin(Config.endOffset));

            if (time > tEariesst && time <= tLatest)
            {
                p.endB = true;
                p.activeCount += 1;
            }

            //     C班上班啊范围
            tEariesst = e.earliestBeginC.AddHours(-(ParseHour(Config.beginOffset)));
            tEariesst = tEariesst.AddMinutes(-(ParseMin(Config.beginOffset)));

            tLatest = e.latestBeginC.AddHours(ParseHour(Config.beginOffset));
            tLatest = tLatest.AddMinutes(ParseMin(Config.beginOffset));

            if (time > tEariesst && time <= tLatest)
            {
                p.startC = true;
                p.activeCount += 1;
            }

            //     C班下班范围
            tEariesst = e.earliestEndC.AddHours(-(ParseHour(Config.endOffset)));
            tEariesst = tEariesst.AddMinutes(-(ParseMin(Config.endOffset)));

            tLatest = e.latestEndC.AddHours(ParseHour(Config.endOffset));
            tLatest = tLatest.AddMinutes(ParseMin(Config.endOffset));

            if (time > tEariesst && time <= tLatest)
            {
                p.endC = true;
                p.activeCount += 1;
            }

            //     D班上班啊范围
            tEariesst = e.earliestBeginD.AddHours(-(ParseHour(Config.beginOffset)));
            tEariesst = tEariesst.AddMinutes(-(ParseMin(Config.beginOffset)));

            tLatest = e.latestBeginD.AddHours(ParseHour(Config.beginOffset));
            tLatest = tLatest.AddMinutes(ParseMin(Config.beginOffset));

            if (time > tEariesst && time <= tLatest)
            {
                p.startD = true;
                p.activeCount += 1;
            }

            //     D班下班范围
            tEariesst = e.earliestEndD.AddHours(-(ParseHour(Config.endOffset)));
            tEariesst = tEariesst.AddMinutes(-(ParseMin(Config.endOffset)));

            //tLatest = e.latestEndD.AddHours(ParseHour(Config.endOffset));
            //tLatest = tLatest.AddMinutes(ParseMin(Config.endOffset));
            tLatest = e.latestEndD.AddHours(8);  //行政班可能加班到2点

            if (time > tEariesst && time <= tLatest)
            {
                p.endD = true;
                p.activeCount += 1;
            }
        }

        public bool IsAD(string name)       //判断是否是行政员工
        {
            //string[] ADlist = {
            //    "曹可生",
            //    "孙龙起",
            //    "黄水珍",
            //    "尹绍翠",
            //    "李阳洋",
            //    "张文东",
            //    "霍小碧",
            //    "潘雨珊",
            //    "刘炬滨",
            //    "朱云龙",
            //};

            foreach (string v in Config.administrativeStaffList)
            {
                if (v == name)
                {
                    return true;
                }
            }

            return false;
        }

        //获取对应的班次
        public MatchingStatus GetReverseSchedule(ProbableScheduling p, CurrentWorkingHours e, DateTime time, bool ADFlag)
        {
            if (ADFlag == true)
            {
                if (p.startD == true)
                {
                    return MatchingStatus.UnMatchedEndD;
                }
                else if (p.endD == true)
                {
                    return MatchingStatus.UnMatchedBeginD;
                }

                return MatchingStatus.Unknow;
            }


            if (p.startA == true || p.startB == true || p.startC == true )     //上班优先
            {
                TimeSpan tsTime = new TimeSpan(time.Ticks);

                int[] offset = { 99999, 99999, 99999, 99999, 99999, 99999};

                if (p.startA == true)
                {
                    TimeSpan tsEBA = new TimeSpan(e.earliestBeginA.Ticks);
                    TimeSpan tsE = tsTime.Subtract(tsEBA);
                    offset[0] = int.Parse(tsE.Duration().TotalSeconds.ToString());

                    TimeSpan tsLBA = new TimeSpan(e.latestBeginA.Ticks);
                    TimeSpan tsL = tsTime.Subtract(tsLBA);
                    offset[1] = int.Parse(tsL.Duration().TotalSeconds.ToString());
                }

                if (p.startB == true)
                {
                    TimeSpan tsEBB = new TimeSpan(e.earliestBeginB.Ticks);
                    TimeSpan tsE = tsTime.Subtract(tsEBB);
                    offset[2] = int.Parse(tsE.Duration().TotalSeconds.ToString());

                    TimeSpan tsLBB = new TimeSpan(e.latestBeginB.Ticks);
                    TimeSpan tsL = tsTime.Subtract(tsLBB);
                    offset[3] = int.Parse(tsL.Duration().TotalSeconds.ToString());
                }

                if (p.startC == true)
                {
                    TimeSpan tsEBC = new TimeSpan(e.earliestBeginC.Ticks);
                    TimeSpan tsE = tsTime.Subtract(tsEBC);
                    offset[4] = int.Parse(tsE.Duration().TotalSeconds.ToString());

                    TimeSpan tsLBC = new TimeSpan(e.latestBeginC.Ticks);
                    TimeSpan tsL = tsTime.Subtract(tsLBC);
                    offset[5] = int.Parse(tsL.Duration().TotalSeconds.ToString());
                }
                
                for(int i = 0; i < offset.Length; i++)
                {
                    if (offset[i] == offset.Min() && offset.Min() != 99999)      //找到距离原始上班班次最近的
                    {
                        if (i == 0 || i == 1)
                        { 
                            //需返回相对应的班次
                            return MatchingStatus.UnMatchedEndA;
                        }
                        else if (i == 2 || i == 3)
                        {
                            return MatchingStatus.UnMatchedEndB;
                        }
                        else if (i == 4 || i == 5)
                        {
                            return MatchingStatus.UnMatchedEndC;
                        }

                        break;
                    }
                }

                return MatchingStatus.Unknow;
            }
            else if (p.endA == true || p.endB == true || p.endC == true)
            {
                TimeSpan tsTime = new TimeSpan(time.Ticks);

                int[] offset = { 99999, 99999, 99999, 99999, 99999, 99999};

                if (p.endA == true)
                {
                    TimeSpan tsEEA = new TimeSpan(e.earliestEndA.Ticks);
                    TimeSpan tsE = tsTime.Subtract(tsEEA);
                    offset[0] = int.Parse(tsE.Duration().TotalSeconds.ToString());

                    TimeSpan tsLEA = new TimeSpan(e.latestEndA.Ticks);
                    TimeSpan tsL = tsTime.Subtract(tsLEA);
                    offset[1] = int.Parse(tsL.Duration().TotalSeconds.ToString());
                }

                if (p.endB == true)
                {
                    TimeSpan tsEEB = new TimeSpan(e.earliestEndB.Ticks);
                    TimeSpan tsE = tsTime.Subtract(tsEEB);
                    offset[2] = int.Parse(tsE.Duration().TotalSeconds.ToString());

                    TimeSpan tsLEB = new TimeSpan(e.latestEndB.Ticks);
                    TimeSpan tsL = tsTime.Subtract(tsLEB);
                    offset[3] = int.Parse(tsL.Duration().TotalSeconds.ToString());
                }

                if (p.endC == true)
                {
                    TimeSpan tsEEC = new TimeSpan(e.earliestEndC.Ticks);
                    TimeSpan tsE = tsTime.Subtract(tsEEC);
                    offset[4] = int.Parse(tsE.Duration().TotalSeconds.ToString());

                    TimeSpan tsLEC = new TimeSpan(e.latestEndC.Ticks);
                    TimeSpan tsL = tsTime.Subtract(tsLEC);
                    offset[5] = int.Parse(tsL.Duration().TotalSeconds.ToString());
                }

                for (int i = 0; i < offset.Length; i++)
                {
                    if (offset[i] == offset.Min() && offset.Min() != 99999)      //找到距离原始上班班次最近的
                    {
                        if (i == 0 || i == 1)
                        {
                            return MatchingStatus.UnMatchedBeginA;
                        }
                        else if (i == 2 || i == 3)
                        {
                            return MatchingStatus.UnMatchedBeginB;
                        }
                        else if (i == 4 || i == 5)
                        {
                            return MatchingStatus.UnMatchedBeginC;
                        }

                        break;
                    }
                }

                return MatchingStatus.Unknow;
            }
            else
            {
                return MatchingStatus.Unknow;
            }
        }

        public MatchingStatus MatchingDetect(Employee e, DateTime time1, DateTime time2)
        {
            string tmp = time1.ToString().Substring(0, time1.ToString().IndexOf(" "));  //找到年月日

            CurrentWorkingHours currentWorkingHours = new CurrentWorkingHours();
            ProbableScheduling probableScheduling1 = new ProbableScheduling();
            ProbableScheduling probableScheduling2 = new ProbableScheduling();

            //实例化当天的上下班时间表
            currentWorkingHours.earliestBeginA = Convert.ToDateTime(tmp + " " + Config.earliestBeginA);
            currentWorkingHours.latestBeginA = Convert.ToDateTime(tmp + " " + Config.latestBeginA);
            currentWorkingHours.earliestEndA = Convert.ToDateTime(tmp + " " + Config.earliestEndA);
            currentWorkingHours.latestEndA = Convert.ToDateTime(tmp + " " + Config.latestEndA);

            currentWorkingHours.earliestBeginB = Convert.ToDateTime(tmp + " " + Config.earliestBeginB);
            currentWorkingHours.latestBeginB = Convert.ToDateTime(tmp + " " + Config.latestBeginB);
            currentWorkingHours.earliestEndB = Convert.ToDateTime(tmp + " " + Config.earliestEndB);
            currentWorkingHours.latestEndB = Convert.ToDateTime(tmp + " " + Config.latestEndB);
            currentWorkingHours.latestEndB = currentWorkingHours.latestEndB.AddDays(1);

            currentWorkingHours.earliestBeginC = Convert.ToDateTime(tmp + " " + Config.earliestBeginC);
            currentWorkingHours.latestBeginC = Convert.ToDateTime(tmp + " " + Config.latestBeginC);
            currentWorkingHours.earliestEndC = Convert.ToDateTime(tmp + " " + Config.earliestEndC);
            currentWorkingHours.earliestEndC = currentWorkingHours.earliestEndC.AddDays(1);
            currentWorkingHours.latestEndC = Convert.ToDateTime(tmp + " " + Config.latestEndC);
            currentWorkingHours.latestEndC = currentWorkingHours.latestEndC.AddDays(1);

            currentWorkingHours.earliestBeginD = Convert.ToDateTime(tmp + " " + Config.earliestBeginD);
            currentWorkingHours.latestBeginD = Convert.ToDateTime(tmp + " " + Config.latestBeginD);
            currentWorkingHours.earliestEndD = Convert.ToDateTime(tmp + " " + Config.earliestEndD);
            currentWorkingHours.latestEndD = Convert.ToDateTime(tmp + " " + Config.latestEndD);

            //解析目标时间可能的班次到到probableScheduling

            GetNearlyScheduling(probableScheduling1, currentWorkingHours, time1);
            GetNearlyScheduling(probableScheduling2, currentWorkingHours, time2);

            //当t1和t2相差超过一天的时候，可能出现probableScheduling2.activeCount = 0
            //当为某人的第一条打卡数据时，可能出现"2021/7/1  1:03:00"这种不在任何范围，activeCount1和activeCount2都为0的情况
            if (probableScheduling1.activeCount == 0 || probableScheduling2.activeCount == 0)
            {
                
                if (probableScheduling1.activeCount == 0)       //第一条数据不在班次时间范围，未未知数据
                {
                    return MatchingStatus.Unknow;
                }
                else  //2=0,1!=0 找对应的1的不匹配数据
                {
                    return GetReverseSchedule(probableScheduling1, currentWorkingHours, time1, IsAD(e.name));
                }
            }

            //有配对的记录
            if (IsAD(e.name))       //行政
            {
                if (probableScheduling1.startD && probableScheduling2.endD)
                {
                    return MatchingStatus.Matched;
                }
                else
                {
                    return GetReverseSchedule(probableScheduling1, currentWorkingHours, time1, true);
                }
            }
            else  //abc班
            {
                if ((probableScheduling1.startA && probableScheduling2.endA) ||
                    (probableScheduling1.startB && probableScheduling2.endB) ||
                    (probableScheduling1.startC && probableScheduling2.endC))
                {
                    return MatchingStatus.Matched;
                }
                else
                {
                    return GetReverseSchedule(probableScheduling1, currentWorkingHours, time1, false);
                }
            }
        }

        public MatchingStatus MatchingDetectLast(Employee e, DateTime time)
        {
            string tmp = time.ToString().Substring(0, time.ToString().IndexOf(" "));  //找到年月日

            CurrentWorkingHours currentWorkingHours = new CurrentWorkingHours();
            ProbableScheduling probableScheduling = new ProbableScheduling();

            //实例化当天的上下班时间表
            currentWorkingHours.earliestBeginA = Convert.ToDateTime(tmp + " " + Config.earliestBeginA);
            currentWorkingHours.latestBeginA = Convert.ToDateTime(tmp + " " + Config.latestBeginA);
            currentWorkingHours.earliestEndA = Convert.ToDateTime(tmp + " " + Config.earliestEndA);
            currentWorkingHours.latestEndA = Convert.ToDateTime(tmp + " " + Config.latestEndA);

            currentWorkingHours.earliestBeginB = Convert.ToDateTime(tmp + " " + Config.earliestBeginB);
            currentWorkingHours.latestBeginB = Convert.ToDateTime(tmp + " " + Config.latestBeginB);
            currentWorkingHours.earliestEndB = Convert.ToDateTime(tmp + " " + Config.earliestEndB);
            currentWorkingHours.latestEndB = Convert.ToDateTime(tmp + " " + Config.latestEndB);
            currentWorkingHours.latestEndB = currentWorkingHours.latestEndB.AddDays(1);

            currentWorkingHours.earliestBeginC = Convert.ToDateTime(tmp + " " + Config.earliestBeginC);
            currentWorkingHours.latestBeginC = Convert.ToDateTime(tmp + " " + Config.latestBeginC);
            currentWorkingHours.earliestEndC = Convert.ToDateTime(tmp + " " + Config.earliestEndC);
            currentWorkingHours.earliestEndC = currentWorkingHours.earliestEndC.AddDays(1);
            currentWorkingHours.latestEndC = Convert.ToDateTime(tmp + " " + Config.latestEndC);
            currentWorkingHours.latestEndC = currentWorkingHours.latestEndC.AddDays(1);

            currentWorkingHours.earliestBeginD = Convert.ToDateTime(tmp + " " + Config.earliestBeginD);
            currentWorkingHours.latestBeginD = Convert.ToDateTime(tmp + " " + Config.latestBeginD);
            currentWorkingHours.earliestEndD = Convert.ToDateTime(tmp + " " + Config.earliestEndD);
            currentWorkingHours.latestEndD = Convert.ToDateTime(tmp + " " + Config.latestEndD);

            //解析目标时间可能的班次到到probableScheduling
            GetNearlyScheduling(probableScheduling, currentWorkingHours, time);

            if (IsAD(e.name))
            {
                if (probableScheduling.startD == true)
                {
                    return MatchingStatus.Matched;  //最后一条为上班记录，则可能下班记录在下个月的表，判定为正常
                }
                else if (probableScheduling.endD == true)   //确实上班记录
                {
                    return MatchingStatus.UnMatchedBeginD;
                }
                else
                {
                    return MatchingStatus.Unknow;
                }
            }
            else
            {
                if (probableScheduling.startA == true || probableScheduling.startB == true || probableScheduling.startC == true)
                {
                    return MatchingStatus.Matched;
                }
                else if (probableScheduling.endA == true || probableScheduling.endB == true || probableScheduling.endC == true)
                {
                    return GetReverseSchedule(probableScheduling, currentWorkingHours, time, false);
                }
                else
                {
                    return MatchingStatus.Unknow;
                }
            }

        }

        //初始化配置，导入班次时间表和行政人员名单
        public void InitConfiguration()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);

            Config.checkInFileName = config.AppSettings.Settings["考勤工作表"].Value.ToString();
            Config.exceptionFileName = config.AppSettings.Settings["异常工作表"].Value.ToString();

            Config.earliestBeginA = config.AppSettings.Settings["earliestBeginA"].Value.ToString();
            Config.latestBeginA = config.AppSettings.Settings["latestBeginA"].Value.ToString();
            Config.earliestEndA = config.AppSettings.Settings["earliestEndA"].Value.ToString();
            Config.latestEndA = config.AppSettings.Settings["latestEndA"].Value.ToString();

            Config.earliestBeginB = config.AppSettings.Settings["earliestBeginB"].Value.ToString();
            Config.latestBeginB = config.AppSettings.Settings["latestBeginB"].Value.ToString();
            Config.earliestEndB = config.AppSettings.Settings["earliestEndB"].Value.ToString();
            Config.latestEndB = config.AppSettings.Settings["latestEndB"].Value.ToString();

            Config.earliestBeginC = config.AppSettings.Settings["earliestBeginC"].Value.ToString();
            Config.latestBeginC = config.AppSettings.Settings["latestBeginC"].Value.ToString();
            Config.earliestEndC = config.AppSettings.Settings["earliestEndC"].Value.ToString();
            Config.latestEndC = config.AppSettings.Settings["latestEndC"].Value.ToString();

            Config.earliestBeginD = config.AppSettings.Settings["earliestBeginD"].Value.ToString();
            Config.latestBeginD = config.AppSettings.Settings["latestBeginD"].Value.ToString();
            Config.earliestEndD = config.AppSettings.Settings["earliestEndD"].Value.ToString();
            Config.latestEndD = config.AppSettings.Settings["latestEndD"].Value.ToString();

            Config.beginOffset = config.AppSettings.Settings["beginOffset"].Value.ToString();
            Config.endOffset = config.AppSettings.Settings["endOffset"].Value.ToString();

            string[] list = config.AppSettings.Settings["行政班名单"].Value.ToString().Split(new char[2] { ',', '，' });

            if (list.Length == 0)
            {
                MessageBox.Show("行政班名单出错！");
                System.Windows.Forms.Application.ExitThread();
                return;
            }

            Config.administrativeStaffList = list;
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.ExitThread();
        }
    } 
}


//配置文件
public static class Config
{
    //打卡记录文件名
    public static string checkInFileName;

    //导出异常文件名
    public static string exceptionFileName;

    //行政人员名单
    public static string[] administrativeStaffList;

    //早班
    public static string earliestBeginA = "4:30";
    public static string latestBeginA = "6:20";
    public static string earliestEndA = "14:00";
    public static string latestEndA = "17:00";

    //中班
    public static string earliestBeginB = "12:00";
    public static string latestBeginB = "14:00";
    public static string earliestEndB = "22:00";
    public static string latestEndB = "2:00";

    //晚班
    public static string earliestBeginC = "16:30";
    public static string latestBeginC = "21:00";
    public static string earliestEndC = "1:00";
    public static string latestEndC = "5:00";

    //正常行政班
    public static string earliestBeginD = "7:30";
    public static string latestBeginD = "8:30";
    public static string earliestEndD = "18:00";
    public static string latestEndD = "18:00";

    //上下班容错时间
    public static string beginOffset = "1:30";
    public static string endOffset = "4:00";
}


//班次工作时间表
//public static class WorkingHours
//{
//    //早班
//    public static string earliestBeginA = "4:30";
//    public static string latestBeginA = "6:20";
//    public static string earliestEndA = "14:00";
//    public static string latestEndA = "17:00";

//    //中班
//    public static string earliestBeginB = "12:00";
//    public static string latestBeginB = "14:00";
//    public static string earliestEndB = "22:00";
//    public static string latestEndB = "2:00";

//    //晚班
//    public static string earliestBeginC = "16:30";
//    public static string latestBeginC = "21:00";
//    public static string earliestEndC = "1:00";
//    public static string latestEndC = "5:00";

//    //正常行政班
//    public static string earliestBeginD = "7:30";
//    public static string latestBeginD = "8:30";
//    public static string earliestEndD = "18:00";
//    public static string latestEndD = "18:00";

//    //上下班容错时间
//    public static string beginOffset = "1:30";
//    public static string endOffset = "4:00";
//}

//当天工作时间表示例化，加入了年月日
public class CurrentWorkingHours
{
    public DateTime earliestBeginA;
    public DateTime latestBeginA;
    public DateTime earliestEndA;
    public DateTime latestEndA;

    public DateTime earliestBeginB;
    public DateTime latestBeginB;
    public DateTime earliestEndB;
    public DateTime latestEndB;

    public DateTime earliestBeginC;
    public DateTime latestBeginC;
    public DateTime earliestEndC;
    public DateTime latestEndC;

    public DateTime earliestBeginD;
    public DateTime latestBeginD;
    public DateTime earliestEndD;
    public DateTime latestEndD;
}

public class ProbableScheduling
{
    public bool startA = false;
    public bool startB = false;
    public bool startC = false;
    public bool startD = false;

    public bool endA = false;
    public bool endB = false;
    public bool endC = false;
    public bool endD = false;

    public int activeCount = 0;
}

public class Employee
{
    public string name;
    public ArrayList checkInTime = new ArrayList();
    public ArrayList CheckInException = new ArrayList();
    public ArrayList CheckInExceptionComments = new ArrayList();
}

public class Employees
{
    public Employee[] em = new Employee[200];
}

public enum MatchingStatus
{
    Matched = 0,
    UnMatchedBeginA,
    UnMatchedEndA,
    UnMatchedBeginB,
    UnMatchedEndB,
    UnMatchedBeginC,
    UnMatchedEndC,
    UnMatchedBeginD,
    UnMatchedEndD,
    Repeat,
    Unknow,
}