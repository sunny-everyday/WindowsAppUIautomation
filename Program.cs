//#define IOT
using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation.Provider;
using System.Windows.Automation.Text;
using System.Windows.Automation;
using System.Reflection;
using MSExcel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Core;
using System.IO;

namespace UIAutomationTest
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("\nBegin WinForm UIAutomation test run\n");

                //自动化根元素
                AutomationElement aeDeskTop = AutomationElement.RootElement;
                if (null == aeDeskTop)
                {
                    Console.WriteLine("DeskTop get fail");
                }

#if BYEXE
                //启动被测试的程序
                Process p = Process.Start(@"D:\Debug(1)\IoTPlatform.exe");
                if (null == p)
                {
                    Console.WriteLine("Process get fail");
                }
                //根据执行程序名获取进程
                Process[]  p2 = new Process[2];
                if (null == Process.GetProcessesByName("IoTPlatform.ext"))
                {
                    Console.WriteLine("Process get fail");
                }
                else
                {
                    p2 = Process.GetProcessesByName("IoTPlatform");
                    
                    {
                        Console.WriteLine("Process get OK");
                    }
                    
                }

                
                Thread.Sleep(10000);
                AutomationElement aeForm = AutomationElement.FromHandle(p2[0].MainWindowHandle);
                //获得对主窗体对象的引用
                if (null == aeForm)
                {
                    Console.WriteLine("Can not find the WinFormTest from.");
                }

                Console.WriteLine("Finding all user controls");
                //找到第一次出现的Button控件
                AutomationElement aeButton = aeForm.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.AutomationIdProperty, "BtnLogin"));

                //找到所有的TextBox控件
                AutomationElementCollection aeAllTextBoxes = aeForm.FindAll(TreeScope.Children,
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));

                //找到所有的下拉框控件
                AutomationElementCollection aeComboBox = aeForm.FindAll(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox));

                // 控件初始化的顺序是先初始化后添加到控件
                // this.Controls.Add(this.textBox3);                  
                // this.Controls.Add(this.textBox2);
                // this.Controls.Add(this.textBox1);

                AutomationElement aeTextBox1 = aeAllTextBoxes[0];
                AutomationElement aeTextBox2 = aeAllTextBoxes[1];
                AutomationElement aeTextComboBox = aeComboBox[0];

               //Console.WriteLine("Settiing user");
                //通过ValuePattern设置TextBox1的值
               ValuePattern vpTextBox1 = (ValuePattern)aeTextBox1.GetCurrentPattern(ValuePattern.Pattern);
               vpTextBox1.SetValue("zhangyz5");
               //Console.WriteLine("Settiing input user");
               //通过ValuePattern设置TextBox2的值
               ValuePattern vpTextBox2 = (ValuePattern)aeTextBox2.GetCurrentPattern(ValuePattern.Pattern);
               vpTextBox2.SetValue("jsepc0730@!");

                //通过ValuePattern设置TextBox3的值
               ValuePattern vpTextBox3 = (ValuePattern)aeTextComboBox.GetCurrentPattern(ValuePattern.Pattern);
               vpTextBox3.SetValue("正式环境");
               Thread.Sleep(1500);
                Console.WriteLine("Clickinig on login Button.");
                //通过InvokePattern模拟点击按钮
                InvokePattern ipClickButton1 = (InvokePattern)aeButton.GetCurrentPattern(InvokePattern.Pattern);
                ipClickButton1.Invoke();

                //实现关闭被测试程序
                //WindowPattern wpCloseForm = (WindowPattern)aeForm.GetCurrentPattern(WindowPattern.Pattern);
                //wpCloseForm.Close();

                //Console.WriteLine("\nEnd test run\n");
#endif
#if DESKTOP
                //根据controltype获取到服务平台
                AutomationElement aeForm = aeDeskTop.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.NameProperty, "输变电物联网服务平台"));
                if (null == aeForm)
                {
                    Console.WriteLine("aeForm get fail");
                    return;
                }
#endif

#if IOT
                //根据执行程序名获取进程
                Process[] p2 = new Process[2];
                p2 = Process.GetProcessesByName("IoTPlatform");
                if (null == p2[0])
                {
                    Console.WriteLine("Process get fail");
                    return;
                }
                else
                {                 
                    Console.WriteLine("Process get OK");         
                }
                Console.WriteLine("Process 1");
                AutomationElement aeForm = AutomationElement.FromHandle(p2[0].MainWindowHandle);
                //获得对主窗体对象的引用
                if (null == aeForm)
                {
                    Console.WriteLine("Can not find the WinFormTest from.");
                    return;
                }
                Console.WriteLine("Process 2");
                //获取子窗口
                AutomationElement aeTabControl = aeForm.FindFirst(TreeScope.Children,
                 new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab));
                if (null == aeTabControl)
                {
                    Console.WriteLine("aeTabControl get fail");
                    return;
                }
                Console.WriteLine("Process 3");
                //获取输电可视化窗口
                AutomationElement aeTabItemControl = aeTabControl.FindFirst(TreeScope.Children,
                 new PropertyCondition(AutomationElement.NameProperty, "输电线路可视化"));
                if (null == aeTabItemControl)
                {
                    Console.WriteLine("aeTabItemControl get fail");
                    return;
                }
                Console.WriteLine("Process 4");
                //获取自定义窗口 
                AutomationElement aeCustomControl = aeTabItemControl.FindFirst(TreeScope.Children,
                 new PropertyCondition(AutomationElement.ClassNameProperty, "OutsidePage"));
                if (null == aeCustomControl)
                {
                    Console.WriteLine("aeCustomControl get fail");
                    return;
                }
                Console.WriteLine("Process 5");
                //找到详情的Button控件
                AutomationElement aeButton = aeCustomControl.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.HelpTextProperty, "详情"));
                if (null == aeButton)
                {
                    Console.WriteLine("aeButton get fail");
                    return;
                }
                Console.WriteLine("Process 6");
#endif
                //从xlsx中获取设备名称
                String[] DeviceName;
                DeviceName = new String[500];
                int DeviceNUM = 0;
                unsafe
                {
                    readxls(DeviceName, &DeviceNUM);
                }
                //依次根据设备名搜索视频信息，记录设备状态
                bool[] DeviceState;
                DeviceState = new bool[500];
                if (DeviceNUM > 500)
                {
                    Console.WriteLine("DeviceNUM bigger than program maximum, please change program.");
                    return;
                }
                for (int Index = 0; Index < DeviceNUM; Index++)
                { 
                    
                }

                //将设备状态记录到xlsx
                writexls(DeviceState, DeviceNUM);

#if IOT
                //通过InvokePattern模拟点击按钮
                InvokePattern ipClickButton = (InvokePattern)aeButton.GetCurrentPattern(InvokePattern.Pattern);
                ipClickButton.Invoke();

                Thread.Sleep(20000);
                
                //找到视频监控界面控件

                AutomationElement aeOutsideDetail = aeCustomControl.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ClassNameProperty, "OutsideDetailes"));
                if (null == aeOutsideDetail)
                {
                    Console.WriteLine("aeOutsideDetail get fail");
                    Thread.Sleep(1000);
                    return;
                }
#endif
#if FORINTURN
                //找到树视图
                AutomationElement aetree = aeOutsideDetail.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ClassNameProperty, "Tree"));
                if (null == aetree)
                {
                    Console.WriteLine("aetree get fail");
                    Thread.Sleep(1000);
                    return;
                }
                //找到ProgressBar
                AutomationElement aeProgressBar = aetree.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ProgressBar));
                if (null == aeProgressBar)
                {
                    Console.WriteLine("aeProgressBar get fail");
                    Thread.Sleep(1000);
                    return;
                }
                //找到tree
                AutomationElement aetree2 = aeProgressBar.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree));
                if (null == aetree2)
                {
                    Console.WriteLine("aetree2 get fail");
                    Thread.Sleep(1000);
                    return;
                }

                //找到treeItem
                AutomationElement aetreeItem = aetree2.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                if (null == aetreeItem)
                {
                    Console.WriteLine("aetreeItem get fail");
                    Thread.Sleep(1000);
                    return;
                }
                //获取所有地市级treeItem
                AutomationElementCollection aeCitytreeItemes = aetreeItem.FindAll(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                if (0 == aeCitytreeItemes.Count)
                {
                    Console.WriteLine("aeCitytreeItem get 0.");
                    Thread.Sleep(1000);
                    return;
                }
                
                int CityNumber = aeCitytreeItemes.Count;
                Console.WriteLine(CityNumber);
                for (int i = 0; i < CityNumber; i++)
                { 
                    //获取地市级公司名
                    AutomationElement aeCityName = aeCitytreeItemes[i].FindFirst(TreeScope.Children,
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
                    if (null != aeCityName)
                    {
                        Console.WriteLine("CityName is ");
                        Console.WriteLine(aeCityName.Current.Name);
                        Thread.Sleep(1000);
 
                    }
                    //展开节点
                    ExpandCollapsePattern ExpandPattern1 = (ExpandCollapsePattern)aeCitytreeItemes[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                    
                    Thread.Sleep(1000);
                    //currentPattern.Collapse();
                    ExpandPattern1.Expand();

                    //区级节点操作
                    Thread.Sleep(10000);
                    //获取该地市区级公司
                    AutomationElementCollection aedistrictItemes = aeCitytreeItemes[i].FindAll(TreeScope.Children,
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                    if (0 == aedistrictItemes.Count)
                    {
                        Console.WriteLine("aedistrictItemes get 0.");
                        Thread.Sleep(1000);
                        return;
                    }
                    int districtNumber = aedistrictItemes.Count;
                    Console.WriteLine(districtNumber);
                    for (int j = 0; j < districtNumber; j++)
                    {
                        //获取区级公司名
                        AutomationElement aedistrictName = aedistrictItemes[j].FindFirst(TreeScope.Children,
                          new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
                        if (null != aedistrictName)
                        {
                            Console.WriteLine("aedistrictName is ");
                            Console.WriteLine(aedistrictName.Current.Name);
                            Thread.Sleep(1000);
                        }
                        //展开节点
                        ExpandCollapsePattern ExpandPattern2 = (ExpandCollapsePattern)aedistrictItemes[j].GetCurrentPattern(ExpandCollapsePattern.Pattern);

                        Thread.Sleep(1000);
                        //currentPattern.Collapse();
                        ExpandPattern2.Expand();

                        //交流电级节点操作
                        Thread.Sleep(10000);
                        AutomationElementCollection aeVoltageItemes = aedistrictItemes[j].FindAll(TreeScope.Children,
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                        if (0 == aeVoltageItemes.Count)
                        {
                            Console.WriteLine("aedistrictItemes get 0.");
                            Thread.Sleep(1000);
                            return;
                        }
                        int VoltageNumber = aeVoltageItemes.Count;
                        Console.WriteLine(VoltageNumber);
                        for (int k = 0; k < VoltageNumber; k++)
                        {
                            //获取交流电压级别名称
                            AutomationElement aeVoltageName = aeVoltageItemes[k].FindFirst(TreeScope.Children,
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
                            if (null != aeVoltageName)
                            {
                                Console.WriteLine("aeVoltageName is ");
                                Console.WriteLine(aeVoltageName.Current.Name);
                                Thread.Sleep(1000);
                            }
                            //展开节点
                            ExpandCollapsePattern ExpandPattern3 = (ExpandCollapsePattern)aeVoltageItemes[k].GetCurrentPattern(ExpandCollapsePattern.Pattern);

                            Thread.Sleep(1000);
                               
                            ExpandPattern3.Expand();

                            //杆塔线路级操作
                            Thread.Sleep(10000);
                            AutomationElementCollection aelineItemes = aeVoltageItemes[k].FindAll(TreeScope.Children,
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                            if (0 == aelineItemes.Count)
                            {
                                Console.WriteLine("aelineItemes get 0.");
                                Thread.Sleep(1000);
                                return;
                            }
                            int lineNumber = aelineItemes.Count;
                            Console.WriteLine(lineNumber);
                            for (int l = 0; l < lineNumber; l++)
                            {
                                //获取杆塔线路名称
                                AutomationElement aelineName = aelineItemes[l].FindFirst(TreeScope.Children,
                                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
                                if (null != aelineName)
                                {
                                    Console.WriteLine("aelineName is ");
                                    Console.WriteLine(aelineName.Current.Name);
                                    Thread.Sleep(1000);
                                }
                                //展开节点
                                ExpandCollapsePattern ExpandPattern4 = (ExpandCollapsePattern)aelineItemes[l].GetCurrentPattern(ExpandCollapsePattern.Pattern);

                                Thread.Sleep(1000);

                                ExpandPattern4.Expand();

                                //杆塔级操作
                                Thread.Sleep(10000);
                                AutomationElementCollection aetowerItemes = aelineItemes[l].FindAll(TreeScope.Children,
                             new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                                if (0 == aetowerItemes.Count)
                                {
                                    Console.WriteLine("aetowerItemes get 0.");
                                    Thread.Sleep(1000);
                                    return;
                                }
                                int towerNumber = aetowerItemes.Count;
                                Console.WriteLine(towerNumber);
                                for (int m = 0; m < towerNumber; m++)
                                {
                                    //获取杆塔名称
                                    AutomationElement towerName = aetowerItemes[m].FindFirst(TreeScope.Children,
                                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
                                    if (null != towerName)
                                    {
                                        Console.WriteLine("towerName is ");
                                        Console.WriteLine(towerName.Current.Name);
                                        Thread.Sleep(1000);
                                    }
                                    //展开节点
                                    ExpandCollapsePattern ExpandPattern5 = (ExpandCollapsePattern)aetowerItemes[m].GetCurrentPattern(ExpandCollapsePattern.Pattern);

                                    Thread.Sleep(1000);

                                    ExpandPattern5.Expand();

                                    //摄像头操作
                                    Console.WriteLine("to be done ");

                                    //折叠节点
                                    ExpandPattern5.Collapse();
                                }

                                //折叠节点
                                ExpandPattern4.Collapse();
                            }


                            //折叠节点
                            ExpandPattern3.Collapse();
                        }

                        //折叠节点
                        ExpandPattern2.Collapse();
                    }
                    //currentPattern.Collapse();
                    ExpandPattern1.Collapse();

                }
#endif

                Console.WriteLine("Did not find it.");
                Console.WriteLine("Test scenario: *FAIL*");
                

                Console.WriteLine("wait for long time.");
                Thread.Sleep(100000);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex.Message);
            }
        }
        unsafe public static bool readxls(string[] DeviceName, int* DeviceNUM)
        {
            string strDir = Directory.GetCurrentDirectory();

            string fileName = strDir + @"\博瑞思运维设备.xlsx";

            MSExcel.Application excelApp= new MSExcel.Application();

            excelApp.Visible = true;//是打开可见

            MSExcel.Workbooks wbks = excelApp.Workbooks;

            MSExcel._Workbook wbk = wbks.Add(fileName);
 

            object Nothing = Missing.Value;

            MSExcel._Worksheet whs;// = wbk.Sheets.Add(Nothing, Nothing, Nothing, Nothing);

            whs = wbk.Sheets[1];//获取第一张工作表

            whs.Activate();
            //取得总记录行数    (包括标题列)

            int rowsint = whs.UsedRange.Cells.Rows.Count; //得到行数

            int columnsint = whs.UsedRange.Cells.Columns.Count;//得到列数



            for (int i = 2; i <= rowsint; i++)
            {
                
                //((Range)worksheet.Cells[1, i + 1]).HorizontalAlignment = XlVAlign.xlVAlignCenter;
                MSExcel.Range rang = (MSExcel.Range)whs.Cells[i, 2];//单元格B2

                DeviceName[i - 2] = rang.Text;//该单元格文本
               
                //whs.Cells[i, 6] = "在线";
                * DeviceNUM += 1;
            }


            wbk.Close();//关闭文档

            wbks.Close();//关闭工作簿

            excelApp.Quit();//关闭excel应用程序

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);//释放excel进程

            excelApp = null;
            return true;

        }
        public static bool writexls(bool[] Devicestate, int DeviceNUM)
        {
            string strDir = Directory.GetCurrentDirectory();

            string fileName = strDir + @"\博瑞思运维设备.xlsx";

            MSExcel.Application excelApp = new MSExcel.Application();

            excelApp.Visible = true;//是打开可见

            MSExcel.Workbooks wbks = excelApp.Workbooks;

            MSExcel._Workbook wbk = wbks.Add(fileName);


            object Nothing = Missing.Value;

            MSExcel._Worksheet whs;// = wbk.Sheets.Add(Nothing, Nothing, Nothing, Nothing);

            whs = wbk.Sheets[1];//获取第一张工作表

            whs.Activate();
            //取得总记录行数    (包括标题列)
            whs.Cells[1, 6] = "设备在线状态";

            for (int i = 2; i < DeviceNUM + 2; i++)
            {

                if (Devicestate[i - 2] == true)
                {
                    whs.Cells[i, 6] = "在线";
                }
                else 
                {
                    whs.Cells[i, 6] = "离线";
                }
               
            }

            excelApp.DisplayAlerts = false;//不弹出是否保存的对话框

            wbk.SaveCopyAs(strDir + @"\博瑞思运维设备_检查结果.xlsx");

            wbk.Close();//关闭文档

            wbks.Close();//关闭工作簿

            excelApp.Quit();//关闭excel应用程序

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);//释放excel进程

            excelApp = null;
            return true;

        }
    }
}