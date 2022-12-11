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
                IXLWorkbook workbook = new XLWorkbook(ofd.FileName);

                bool flag = false;
                foreach (IXLWorksheet w in workbook.Worksheets)
                {
                    if (w.Name == Config.checkInFileName)
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

                IXLWorksheet sheet = workbook.Worksheet(Config.checkInFileName);

                Employees eplys = new Employees();

                foreach (IXLRow r in sheet.Rows())      //获取所有打卡记录
                {
                    if (r.Cell(1).Value.ToString() == "员工姓名")
                    {
                        continue;
                    }

                    bool exsitFlag = false;
                    for (int i = 0; i < Config.maxEmployeeNum; i++)
                    {
                        if (eplys.employee[i] != null)
                        {
                            if (eplys.employee[i].name == r.Cell(1).Value.ToString())       //找到此人的记录
                            {
                                exsitFlag = true;
                                
                                if (r.Cell(2).Value.ToString() == "签到")
                                {
                                    eplys.employee[i].addCheckInTime("签到", Convert.ToDateTime(r.Cell(3).Value.ToString()));
                                }
                                else
                                {
                                    eplys.employee[i].addCheckInTime("签退", Convert.ToDateTime(r.Cell(4).Value.ToString()));
                                }

                                break;
                            }
                        }
                    }

                    if (exsitFlag == false)     //未找到此人的记录，新增加一个此人的
                    {
                        for (int i = 0; i < Config.maxEmployeeNum; i++)
                        {
                            if (eplys.employee[i] == null)
                            {
                                eplys.employee[i] = new Employee();
                                eplys.employee[i].name = r.Cell(1).Value.ToString();

                                if (Config.groupCList.Contains(eplys.employee[i].name))
                                {
                                    eplys.employee[i].group = "C";
                                }
                                else if (Config.groupDList.Contains(eplys.employee[i].name))
                                {
                                    eplys.employee[i].group = "D";
                                }
                                else if (Config.groupEList.Contains(eplys.employee[i].name))
                                {
                                    eplys.employee[i].group = "E";
                                }
                                else if (Config.groupFList.Contains(eplys.employee[i].name))
                                {
                                    eplys.employee[i].group = "F";
                                }



                                if (r.Cell(2).Value.ToString() == "签到")
                                {
                                    eplys.employee[i].addCheckInTime("签到", Convert.ToDateTime(r.Cell(3).Value.ToString()));
                                }
                                else
                                {
                                    eplys.employee[i].addCheckInTime("签退", Convert.ToDateTime(r.Cell(4).Value.ToString()));
                                }



                                break;
                            }
                        }
                    }
                }

                //分析异常数据
                for (int i = 0; i < eplys.employee.Count(); i++)
                {
                    if (eplys.employee[i] != null)
                    { 
                        bool currentMatched = false, nextMatched = false;
                        string currentComment = "", nextComment = "";

                        for (int j = 0; j < eplys.employee[i].checkInfo.Count() - 1;j++)      //需要读取j+1，末尾控制，Count是全索引max
                        {
                            bool isOverTime = false;

                            if (j == 0 && eplys.employee[i].checkInfo[j].checkType == "签退") //首记录为发签退，直接跳过
                            {
                                continue;
                            }

                            if (eplys.employee[i].checkInfo[j] != null)                        
                            {
 
                                DateTime tCurrent = Convert.ToDateTime(eplys.employee[i].checkInfo[j].checkInTime);
                                DateTime tNext;

                                if (eplys.employee[i].checkInfo[j + 1] == null)  //达到最末尾的数据,结构数组Count全访问
                                {
                                    //需要单独处理最后一个元素
                                }
                                else
                                {
                                    tNext = Convert.ToDateTime(eplys.employee[i].checkInfo[j + 1].checkInTime);
                                    if (j > 0)  //检查前后间隔是否大于10分钟,加班打卡
                                    {
                                        DateTime tPrev = Convert.ToDateTime(eplys.employee[i].checkInfo[j - 1].checkInTime);

                                        //和上一次数据间隔小于10分钟为重复签到
                                        TimeSpan tsCurrent = new TimeSpan(tCurrent.Ticks);
                                        TimeSpan tsPrev = new TimeSpan(tPrev.Ticks);

                                        double differPrev = tsCurrent.Subtract(tsPrev).Duration().TotalMinutes;

                                        if (differPrev <= 20 && eplys.employee[i].checkInfo[j].checkType == eplys.employee[i].checkInfo[j - 1].checkType)  //间隔很小，且为同类型打卡则判定为同类型打卡
                                        {
                                            if (Config.isDispRepeat)
                                            {
                                                eplys.employee[i].checkInfo[j].comment = "重复打卡";
                                            }

                                            continue;
                                        }
                                        else if (differPrev < 10 && eplys.employee[i].checkInfo[j - 1].checkType == "签退" && eplys.employee[i].checkInfo[j].checkType == "签到") //间隔很小，前者下班，后者上班，则为加班记录
                                        {
                                            isOverTime = true; //加班
                                        }
                                       
                                    }

                                    //if (Convert.ToDateTime("2022/11/1 6:49:37") == tCurrent)
                                    //{
                                    //    var t = 1;
                                    //}

                                    (currentMatched, nextMatched, currentComment, nextComment)
                                        = MatchingPairs(eplys.employee[i].group, eplys.employee[i].checkInfo[j].checkType, tCurrent, eplys.employee[i].checkInfo[j + 1].checkType, tNext, isOverTime);

                                    if (currentMatched && nextMatched)  //配对成功
                                    {
                                        eplys.employee[i].checkInfo[j].comment = currentComment;
                                        eplys.employee[i].checkInfo[j + 1].comment = nextComment;

                                        j += 1; //配对索引多+1
                                    }
                                    else if (currentMatched && !nextMatched)    //仅当前成功
                                    {
                                        eplys.employee[i].checkInfo[j].comment = currentComment;
                                    }
                                    else      //未有任何匹配成功
                                    {
                                        eplys.employee[i].checkInfo[j].comment = currentComment;
                                    }
                                }
                            }                         
                        }

                    }
                }


                //更新异常数据表
                //删除原来的统计表
                foreach(IXLWorksheet w in workbook.Worksheets)
                {
                    if (w.Name == Config.exceptionFileName)
                    {
                        workbook.Worksheet(Config.exceptionFileName).Delete();
                        break;
                    }
                }
                //重新生成统计表
                IXLWorksheet sheetException = workbook.AddWorksheet(Config.exceptionFileName);
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


                //列宽
                sheetException.Column(3).Width = 8.88;
                sheetException.Column(4).Width = 18.88;
                sheetException.Column(5).Width = 16.5;
                sheetException.Column(6).Width = 16.38;
                sheetException.Column(7).Width = 12.5;
                sheetException.Column(8).Width = 13.25;
                sheetException.Column(9).Width = 8.88;
                sheetException.Column(10).Width = 14.38;
                sheetException.Column(11).Width = 18.5;
                sheetException.Column(12).Width = 18.5;

                //输出异常到excle和form
                for (int i = 0; i < Config.maxEmployeeNum; i++)
                {
                    if (eplys.employee[i] != null)
                    {
                        for (int j = 0; j < eplys.employee[i].checkInfo.Count(); j++)
                        {
                            if (eplys.employee[i].checkInfo[j] != null && eplys.employee[i].checkInfo[j].comment != null && eplys.employee[i].checkInfo[j].comment != "")
                            {

                                sheetException.Cell(lineIndex, 3).Value = eplys.employee[i].name;
                                sheetException.Cell(lineIndex, 5).Value = eplys.employee[i].checkInfo[j].checkInTime;
                                sheetException.Cell(lineIndex, 6).Value = eplys.employee[i].checkInfo[j].comment;

                                //输出到form
                                string nameFix = "         ".Substring(0, (4 - eplys.employee[i].name.Length) * 2);
                                richTextBoxException.Text += nameFix + eplys.employee[i].name + "      ";
                                richTextBoxException.Text += String.Format("{0, -24}{1, -15}\n", eplys.employee[i].checkInfo[j].comment, eplys.employee[i].checkInfo[j].comment);

                                lineIndex++;
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


        DateTime getEariesstTime(DateTime fixedDate, string offset)
        {
            DateTime tmp= fixedDate;
            tmp = tmp.AddHours(-(ParseHour(offset.ToString())));
            tmp = tmp.AddMinutes(-(ParseMin(offset)));

            return tmp;
        }

        DateTime getLatestTime(DateTime fixedDate, string offset)
        {
            DateTime tmp = fixedDate;
            tmp = tmp.AddHours(ParseHour(offset));
            tmp = tmp.AddMinutes(ParseMin(offset));

            return tmp;
        }

        DateTime getBeginFloatTime(DateTime fixedDate, string offset)
        {
            DateTime tmp = fixedDate;
            tmp = tmp.AddHours(ParseHour(offset));
            tmp = tmp.AddMinutes(ParseMin(offset));
            return tmp;
        }

        DateTime getEndFloatTime(DateTime fixedDate, string offset)
        {
            DateTime tmp = fixedDate;
            tmp = tmp.AddHours(-ParseHour(offset));
            tmp = tmp.AddMinutes(-ParseMin(offset));
            return tmp;
        }

        (bool, string) isInRange(CurrentWorkingHours currentWorkingHours, DateTime time, string group, string checkInType, bool isOverTime)
        {
            string tmpGroup = group;

            if (tmpGroup == null || tmpGroup == "")    //AB组不固定，手动区分你
            {
                if (checkInType == "签到")    //签到
                {
                    if (time >= getEariesstTime(currentWorkingHours.beginA, Config.beginOffset) && time < getLatestTime(currentWorkingHours.beginA, Config.beginOffset)) //满足A班范围
                    {
                        tmpGroup = "A";
                    }
                    else if (time >= getEariesstTime(currentWorkingHours.beginB, Config.beginOffset) && time < getLatestTime(currentWorkingHours.beginB, Config.beginOffset)) //满足B班范围
                    {
                        tmpGroup = "B";
                    }
                    else if (time >= getEariesstTime(currentWorkingHours.beginC, Config.beginOffset) && time < getLatestTime(currentWorkingHours.beginC, Config.beginOffset)) //满足C班范围
                    {
                        tmpGroup = "C";
                    }
                    else
                    {
                        return (false, "未识别班次");
                    }
                }
                else    //签退
                {
                    if (time > currentWorkingHours.beginA && time <= getLatestTime(currentWorkingHours.endA, Config.endOffset)) //满足A班范围
                    {
                        tmpGroup = "A";
                    }
                    else if (time > currentWorkingHours.beginB && time <= getLatestTime(currentWorkingHours.endB, Config.endOffset)) //满足B班范围
                    {
                        tmpGroup = "B";
                    }
                    else if (time > currentWorkingHours.beginC && time <= getLatestTime(currentWorkingHours.endC, Config.endOffset)) //满足B班范围
                    {
                        tmpGroup = "C";
                    }
                    else
                    {
                        return (false, "未识别班次");
                    }

                    //Globle.prevGrop = tmpGroup;  //如果非固定班加班，需要根据上一次成功的判断
                }
            }

            if (tmpGroup == "A")
            {
                if (checkInType == "签到")    //签到
                {
                    if (time >= getEariesstTime(currentWorkingHours.beginA, Config.beginOffset) && time < currentWorkingHours.endA) //满足上班范围
                    {
                        //if (time <= currentWorkingHours.beginA)
                        if (time <= getBeginFloatTime(currentWorkingHours.beginA, Config.lateBeginFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "上班迟到");       //上班迟到
                            }

                            return (true, "");       //上班迟到
                        }
                    }
                    
                    return (false, "未知");
                }
                else    //签退
                {
                    if (time > currentWorkingHours.beginA && time <= getLatestTime(currentWorkingHours.endA, Config.endOffset)) //满足下班范围
                    {
                        //if (time >= currentWorkingHours.endA)
                        if (time >= getEndFloatTime(currentWorkingHours.endA, Config.earlyEndFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "下班早退");       //上班迟到
                            }

                            return (true, ""); 
                        }
                    }

                    return (false, "未知");
                }
            }
            else if (tmpGroup == "B")
            {
                if (checkInType == "签到")    //签到
                {
                    if (time >= getEariesstTime(currentWorkingHours.beginB, Config.beginOffset) && time < currentWorkingHours.endB) //满足上班范围
                    {
                        //if (time <= currentWorkingHours.beginB)
                        if (time <= getBeginFloatTime(currentWorkingHours.beginB, Config.lateBeginFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "上班迟到");       //上班迟到
                            }

                            return (true, "");
                        }
                    }

                    return (false, "未知");
                }
                else    //签退
                {
                    if (time > currentWorkingHours.beginB && time <= getLatestTime(currentWorkingHours.endB, Config.endOffset)) //满足下班范围
                    {
                        //if (time >= currentWorkingHours.endB)
                        if (time >= getEndFloatTime(currentWorkingHours.endB, Config.earlyEndFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "下班早退");       //上班迟到
                            }

                            return (true, "");
                        }
                    }

                    return (false, "未知");
                }
            }
            else if (tmpGroup == "C")
            {
                if (checkInType == "签到")    //签到
                {
                    if (time >= getEariesstTime(currentWorkingHours.beginC, Config.beginOffset) && time < currentWorkingHours.endC) //满足上班范围
                    {
                        //if (time <= currentWorkingHours.beginC)
                        if (time <= getBeginFloatTime(currentWorkingHours.beginC, Config.lateBeginFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "上班迟到");       //上班迟到
                            }

                            return (true, "");
                        }
                    }

                    return (false, "未知");
                }
                else    //签退
                {
                    if (time > currentWorkingHours.beginC && time <= getLatestTime(currentWorkingHours.endC, Config.endOffset)) //满足下班范围
                    {
                        //if (time >= currentWorkingHours.endC)
                        if (time >= getEndFloatTime(currentWorkingHours.endC, Config.earlyEndFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "下班早退");       //上班迟到
                            }

                            return (true, "");
                        }
                    }

                    return (false, "未知");
                }
            }
            else if (tmpGroup == "D")
            {
                if (checkInType == "签到")    //签到
                {
                    DateTime tmp = getEariesstTime(currentWorkingHours.beginD, Config.beginOffset);
                    if (time >= getEariesstTime(currentWorkingHours.beginD, Config.beginOffset) && time < currentWorkingHours.endD) //满足上班范围
                    {
                        //if (time <= currentWorkingHours.beginD)
                        if (time <= getBeginFloatTime(currentWorkingHours.beginD, Config.lateBeginFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "上班迟到");       //上班迟到
                            }

                            return (true, "");
                        }
                    }

                    return (false, "未知");
                }
                else    //签退
                {
                    if (time > currentWorkingHours.beginD && time <= getLatestTime(currentWorkingHours.endD, Config.endOffset)) //满足下班范围
                    {
                        //if (time >= currentWorkingHours.endD)
                        if (time >= getEndFloatTime(currentWorkingHours.endD, Config.earlyEndFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "下班早退");       //上班迟到
                            }

                            return (true, "");
                        }
                    }

                    return (false, "未知");
                }
            }
            else if (tmpGroup == "E")
            {
                if (checkInType == "签到")    //签到
                {
                    if (time >= getEariesstTime(currentWorkingHours.beginE, Config.beginOffset) && time < currentWorkingHours.endE) //满足上班范围
                    {
                        //if (time <= currentWorkingHours.beginE)
                        if (time <= getBeginFloatTime(currentWorkingHours.beginE, Config.lateBeginFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "上班迟到");       //上班迟到
                            }

                            return (true, "");
                        }
                    }

                    return (false, "未知");
                }
                else    //签退
                {
                    if (time > currentWorkingHours.beginE && time <= getLatestTime(currentWorkingHours.endE, Config.endOffset)) //满足下班范围
                    {
                        //if (time >= currentWorkingHours.endE
                        if (time >= getEndFloatTime(currentWorkingHours.endE, Config.earlyEndFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (Config.isDispLateEarly)
                            {
                                return (true, "下班早退");       //上班迟到
                            }

                            return (true, "");
                        }
                    }

                    return (false, "未知");
                }
            }
            else if (tmpGroup == "F")
            {
                if (checkInType == "签到")    //签到
                {
                    if (time >= getEariesstTime(currentWorkingHours.beginF, Config.beginOffset) && time < currentWorkingHours.endF) //满足上班范围
                    {
                        //if (time <= currentWorkingHours.beginF)
                        var t = getBeginFloatTime(currentWorkingHours.beginF, Config.lateBeginFloat);
                        if (time <= getBeginFloatTime(currentWorkingHours.beginF, Config.lateBeginFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (isOverTime)        //加班不计迟到
                            {
                                return (true, "");
                            }
                            else
                            {
                                if (Config.isDispLateEarly)
                                {
                                    return (true, "上班迟到");       //上班迟到
                                }

                                return (true, "");
                            }
                            
                        }
                    }

                    return (false, "未知");
                }
                else    //签退
                {
                    if (time > currentWorkingHours.beginF && time <= getLatestTime(currentWorkingHours.endF, Config.endOffset)) //满足下班范围
                    {
                        //if (time >= currentWorkingHours.endF)
                        if (time >= getEndFloatTime(currentWorkingHours.endF, Config.earlyEndFloat))
                        {
                            return (true, "");                 //上班正常
                        }
                        else
                        {
                            if (isOverTime)        //加班不计早退
                            {
                                return (true, "");
                            }
                            else
                            {
                                if (Config.isDispLateEarly)
                                {
                                    return (true, "下班早退");       //上班迟到
                                }

                                return (true, "");
                            }
                        }
                    }

                    return (false, "未知");
                }
            }

            return (false, "未知");
        }

        public (bool currentMatched, bool nextMatched, string currentComment, string nextComment) MatchingPairs
            (string group, string typeCurrent, DateTime timeCurrent, string typeNext, DateTime timeNext, bool isOverTime)
        {
            bool cMathend = false, nMathend = false;
            string cComment = "未知", nComment = "未知";

            if (typeCurrent == null) return (cMathend, nMathend, "记录类型未知", nComment);

            if (typeNext == null) return (cMathend, nMathend, "记录类型未知next", nComment);

            string tmp = timeCurrent.ToString().Substring(0, timeCurrent.ToString().IndexOf(" "));  //找到年月日

            CurrentWorkingHours currentWorkingHours = new CurrentWorkingHours();

            //实例化当天的上下班时间表
            if (!isOverTime)
            {
                currentWorkingHours.beginA = Convert.ToDateTime(tmp + " " + Config.beginA);
                currentWorkingHours.endA = Convert.ToDateTime(tmp + " " + Config.endA);

                currentWorkingHours.beginB = Convert.ToDateTime(tmp + " " + Config.beginB);
                currentWorkingHours.endB = Convert.ToDateTime(tmp + " " + Config.endB);

                currentWorkingHours.beginC = Convert.ToDateTime(tmp + " " + Config.beginC);
                currentWorkingHours.endC = Convert.ToDateTime(tmp + " " + Config.endC);
                currentWorkingHours.endC = currentWorkingHours.endC.AddDays(1);

                currentWorkingHours.beginD = Convert.ToDateTime(tmp + " " + Config.beginD);
                currentWorkingHours.endD = Convert.ToDateTime(tmp + " " + Config.endD);

                currentWorkingHours.beginE = Convert.ToDateTime(tmp + " " + Config.beginE);
                currentWorkingHours.endE = Convert.ToDateTime(tmp + " " + Config.endE);
                currentWorkingHours.endE = currentWorkingHours.endE.AddDays(1);

                currentWorkingHours.beginF = Convert.ToDateTime(tmp + " " + Config.beginF);
                currentWorkingHours.endF = Convert.ToDateTime(tmp + " " + Config.endF);
            }
            else  //加班时间段
            {
                currentWorkingHours.beginA = Convert.ToDateTime(tmp + " " + Config.endA);
                currentWorkingHours.endA = currentWorkingHours.beginA.AddHours(ParseHour(Config.overTimeDuration));
                currentWorkingHours.endA = currentWorkingHours.beginA.AddMinutes(ParseMin(Config.overTimeDuration));

                currentWorkingHours.beginB = Convert.ToDateTime(tmp + " " + Config.endB);
                currentWorkingHours.endB = currentWorkingHours.beginB.AddHours(ParseHour(Config.overTimeDuration));
                currentWorkingHours.endB = currentWorkingHours.beginB.AddMinutes(ParseMin(Config.overTimeDuration));

                currentWorkingHours.beginC = Convert.ToDateTime(tmp + " " + Config.endC);
                currentWorkingHours.endC = currentWorkingHours.beginC.AddHours(ParseHour(Config.overTimeDuration));
                currentWorkingHours.endC= currentWorkingHours.beginC.AddMinutes(ParseMin(Config.overTimeDuration));

                currentWorkingHours.beginD = Convert.ToDateTime(tmp + " " + Config.endD);
                currentWorkingHours.endD = currentWorkingHours.beginD.AddHours(ParseHour(Config.overTimeDuration));
                currentWorkingHours.endD = currentWorkingHours.beginD.AddMinutes(ParseMin(Config.overTimeDuration));

                currentWorkingHours.beginE = Convert.ToDateTime(tmp + " " + Config.endE);
                currentWorkingHours.endE = currentWorkingHours.beginE.AddHours(ParseHour(Config.overTimeDuration));
                currentWorkingHours.endE = currentWorkingHours.beginE.AddMinutes(ParseMin(Config.overTimeDuration));

                currentWorkingHours.beginF = Convert.ToDateTime(tmp + " " + Config.endF);
                currentWorkingHours.endF = currentWorkingHours.beginF.AddHours(ParseHour(Config.overTimeDuration));
                currentWorkingHours.endA = currentWorkingHours.beginF.AddMinutes(ParseMin(Config.overTimeDuration));

            }


            if (typeCurrent == "签到")  //签到
            {
                if (isOverTime)     //加班签到直接判定，不进入后边流程
                {
                    DateTime tmpTime = timeCurrent;
                    tmpTime = tmpTime.AddHours(ParseHour(Config.overTimeDuration) + 6);
                    tmpTime = tmpTime.AddMinutes(ParseMin(Config.overTimeDuration));

                    if (timeNext <= tmpTime)
                    {
                        return (true, true, "", "");
                    }
                    else
                    {
                        return (false, false, "无加班下班数据", "");
                    }
                }

              
                var (isInCurrent, commentCurrent) = isInRange(currentWorkingHours, timeCurrent, group, typeCurrent, isOverTime);

                if (isInCurrent == true)   //在签到范围之内，匹配签退
                {
                    cMathend = true;
                    cComment = commentCurrent;


                    var (isInNext, commentNext) = isInRange(currentWorkingHours, timeNext, group, typeNext, isOverTime);
                    if (isInNext == true)   //同时在签退范围
                    {
                        nMathend = true;
                        nComment = commentNext;
                    }
                    else  //不在签退范围
                    {
                        nMathend = false;
                        cComment = "无下班记录";
                    }

                }
            }
            else  //签退
            {
                var (isInCurrent, commentCurrent) = isInRange(currentWorkingHours, timeCurrent, group, typeCurrent, isOverTime);
                if (isInCurrent == true)   //在签退范围之内，不继续匹配，将下一个记录丢给下一次匹配
                {
                    cMathend = true;
                    cComment = "无上班记录";
                }
                else
                {
                    cMathend = false;
                    cComment = "无上班记录2";
                }

            }

            return (cMathend, nMathend, cComment, nComment);

        }


        //初始化配置，导入班次时间表和行政人员名单
        public void InitConfiguration()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);

            Config.maxEmployeeNum = Convert.ToInt32(config.AppSettings.Settings["员工总数"].Value) + 10;
            Config.maxCheckInfoNum = Convert.ToInt32(config.AppSettings.Settings["单人最多记录条数"].Value) + 10;
            
            Config.checkInFileName = config.AppSettings.Settings["考勤工作表"].Value.ToString();
            Config.exceptionFileName = config.AppSettings.Settings["异常工作表"].Value.ToString();

            if (config.AppSettings.Settings["显示重复签到"].Value.ToString() == "是")
            {
                Config.isDispRepeat = true;
            }

            if (config.AppSettings.Settings["显示迟到早退"].Value.ToString() == "是")
            {
                Config.isDispLateEarly = true;
            }


            string[] listGroupA = config.AppSettings.Settings["A班名单"].Value.ToString().Split(new char[3] { ',', '，', '、' });
            string[] listGroupB = config.AppSettings.Settings["B班名单"].Value.ToString().Split(new char[3] { ',', '，', '、' });
            string[] listGroupC = config.AppSettings.Settings["C班名单"].Value.ToString().Split(new char[3] { ',', '，', '、' });
            string[] listGroupD = config.AppSettings.Settings["D班名单"].Value.ToString().Split(new char[3] { ',', '，', '、' });
            string[] listGroupE = config.AppSettings.Settings["E班名单"].Value.ToString().Split(new char[3] { ',', '，', '、' });
            string[] listGroupF = config.AppSettings.Settings["行政班名单"].Value.ToString().Split(new char[3] { ',', '，', '、' });

            Config.groupAList= listGroupA;
            Config.groupBList= listGroupB;
            Config.groupCList= listGroupC;
            Config.groupDList= listGroupD;
            Config.groupEList= listGroupE;
            Config.groupFList= listGroupF;


            Config.beginA = config.AppSettings.Settings["A班上班时间"].Value.ToString();
            Config.endA = config.AppSettings.Settings["A班下班时间"].Value.ToString();

            Config.beginB = config.AppSettings.Settings["B班上班时间"].Value.ToString();
            Config.endB = config.AppSettings.Settings["B班下班时间"].Value.ToString();

            Config.beginC = config.AppSettings.Settings["C班上班时间"].Value.ToString();
            Config.endC = config.AppSettings.Settings["C班下班时间"].Value.ToString();

            Config.beginD = config.AppSettings.Settings["D班上班时间"].Value.ToString();
            Config.endD = config.AppSettings.Settings["D班下班时间"].Value.ToString();

            Config.beginE = config.AppSettings.Settings["E班上班时间"].Value.ToString();
            Config.endE = config.AppSettings.Settings["E班下班时间"].Value.ToString();

            Config.beginF = config.AppSettings.Settings["行政班上班时间"].Value.ToString();
            Config.endF = config.AppSettings.Settings["行政班下班时间"].Value.ToString();

            Config.beginOffset = config.AppSettings.Settings["上班时间容错区间"].Value.ToString();
            Config.endOffset = config.AppSettings.Settings["下班时间容错区间"].Value.ToString();

            Config.overTimeDuration = config.AppSettings.Settings["最大加班时常"].Value.ToString();

            Config.lateBeginFloat = config.AppSettings.Settings["迟到浮动时间"].Value.ToString();
            Config.earlyEndFloat = config.AppSettings.Settings["早退浮动时间"].Value.ToString();
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
    //最大员工数量
    public static int maxEmployeeNum = 200;

    //单个员工最多记录
    public static int maxCheckInfoNum = 12 * 60;

    //打卡记录文件名
    public static string checkInFileName;

    //导出异常文件名
    public static string exceptionFileName;

    public static bool isDispRepeat = false;

    public static bool isDispLateEarly = false;

    //各个班次人员列表
    public static string[] groupAList;
    public static string[] groupBList;
    public static string[] groupCList;
    public static string[] groupDList;
    public static string[] groupEList;
    public static string[] groupFList;

    //班次
    public static string beginA = "6:30";
    public static string endA = "14:30";

    public static string beginB = "14:30";
    public static string endB = "22:30";

    public static string beginC = "22:30";
    public static string endC = "6:30";

    public static string beginD = "7:00";
    public static string endD = "19:00";

    public static string beginE = "19:00";
    public static string endE = "7:00";

    //F 行政人员名单
    public static string beginF = "8:30";
    public static string endF = "17:30";

    public static string beginOffset = "1:30";
    public static string endOffset = "9:00";

    public static string overTimeDuration = "8:00";

    public static string lateBeginFloat = "0:10";
    public static string earlyEndFloat = "0:10";
}




//当天工作时间表示例化，加入了年月日
public class CurrentWorkingHours
{
    public DateTime beginA;
    public DateTime endA;

    public DateTime beginB;
    public DateTime endB;

    public DateTime beginC;
    public DateTime endC;

    public DateTime beginD;
    public DateTime endD;

    public DateTime beginE;
    public DateTime endE;

    public DateTime beginF;
    public DateTime endF;
}

public class CheckInInfo
{
    public string checkType;
    public DateTime checkInTime;
    public string comment;
}

public class Employee
{
    public string name;
    public string group;

    public ArrayList CheckInExceptionComments = new ArrayList();

    public CheckInInfo[] checkInfo = new CheckInInfo[Config.maxCheckInfoNum];

    public void addCheckInTime(string status, DateTime checkInTime)
    {
        for (int i = 0; i < Config.maxEmployeeNum; i++)
        {
            if (checkInfo[i] == null)
            {
                checkInfo[i] = new CheckInInfo();
                checkInfo[i].checkType = status;
                checkInfo[i].checkInTime = checkInTime;
                break;
            }
        }
    }
}

public class Employees
{
    public Employee[] employee = new Employee[Config.maxEmployeeNum];
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

