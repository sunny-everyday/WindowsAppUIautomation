using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation.Provider;
using System.Windows.Automation.Text;
using System.Windows.Automation;
using System.Windows;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using MSExcel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Core;
using System.IO;


namespace UIAutomationTest
{
    
    class Program
    {
        private static readonly int MOUSEEVENTF_MOVE  = 0x0001;//模拟鼠标移动
        public static readonly int MOUSEEVENTF_LEFTDOWN = 0x0002;//模拟鼠标左键按下
        private static readonly int MOUSEEVENTF_LEFTUP = 0x0004;//模拟鼠标左键抬起
        private readonly int MOUSEEVENTF_ABSOLUTE = 0x8000;//鼠标绝对位置
        private readonly int MOUSEEVENTF_RIGHTDOWN = 0x0008; //模拟鼠标右键按下 
        private readonly int MOUSEEVENTF_RIGHTUP = 0x0010; //模拟鼠标右键抬起 
        private readonly int MOUSEEVENTF_MIDDLEDOWN = 0x0020; //模拟鼠标中键按下 
        private readonly int MOUSEEVENTF_MIDDLEUP = 0x0040;// 模拟鼠标中键抬起 
        
        public struct PONITAPI
        {
            public int x, y;
        };
        public struct DeviceInfo
        {
            public string DeviceName;
            public bool   DeviceState;
        };
        [DllImport("user32.dll")]
        public static extern int GetCursorPos(ref PONITAPI p);

        [DllImport("user32.dll")]
        public static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("\nBegin WinForm UIAutomation test run\n");

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

                //从xlsx中获取设备名称
                String[] DeviceName;
                DeviceName = new String[500];
                int DeviceNUM = 0;
                unsafe
                {
                    readxls(DeviceName, &DeviceNUM);
                }
                if (0 == DeviceNUM)
                {
                    Console.WriteLine("Get NO Device from file.");
                    return;
                }

                //点击详情按键，进入设备信息界面
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
                AutomationElement aeSearch = GetSearchText(aeOutsideDetail);
                if (null == aeSearch)
                {
                    Console.WriteLine("aeSearch get fail");
                    Thread.Sleep(1000);
                    return;
                };
                AutomationElement aeBasicTreeItem = GetBasicTreeItem(aeOutsideDetail);
                //依次根据设备名搜索视频信息，记录设备状态
                bool[] DeviceState;
                DeviceState = new bool[500];
                if (DeviceNUM > 500)
                {
                    Console.WriteLine("DeviceNUM bigger than program maximum, please change program.");
                    return;
                }
                DeviceInfo[] deviceInfo = new DeviceInfo[500];
                int DealDeviceNum = OpenCameraNode(aeBasicTreeItem, deviceInfo, aeOutsideDetail);
                if(DealDeviceNum == 0)
                {
                    Console.WriteLine("DealDeviceNum is zero.");
                    return;
                }
                for (int Index = 0; Index < Math.Min(DeviceNUM, DealDeviceNum); Index++)
                {
                    //搜索设备，查看视频状态, 记录状态
                    if (false == GetDevicestate(DeviceName[Index], deviceInfo, DeviceNUM))
                    {
                        DeviceState[Index] = false;
                        Console.WriteLine("Device Index",Index,"can't find");
                    }
                    else
                    {
                        DeviceState[Index] = true;
                        Console.WriteLine("Device Index", Index, "can find");
                    }
                    
                }

                //将设备状态记录到xlsx
                writexls(DeviceState, DeviceNUM);

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
        public static AutomationElement GetCityNodeWhichexpand(AutomationElement aeOutsideDetail)
        {
            AutomationElement aetreeItem = GetBasicTreeItem(aeOutsideDetail);
            //获取所有地市级treeItem
            AutomationElementCollection aeCitytreeItemes = aetreeItem.FindAll(TreeScope.Children,
              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
            if (0 == aeCitytreeItemes.Count)
            {
                Console.WriteLine("aeCitytreeItem get 0.");
                Thread.Sleep(1000);
                return null;
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
                    //Console.WriteLine("CityName is ");
                    //Console.WriteLine(aeCityName.Current.Name);
                    //Thread.Sleep(1000);
                }
                //展开节点
                ExpandCollapsePattern ExpandPattern1 = (ExpandCollapsePattern)aeCitytreeItemes[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                if (ExpandPattern1.Current.ExpandCollapseState == ExpandCollapseState.Expanded)
                {
                    Console.WriteLine(aeCityName.Current.Name);
                    return aeCitytreeItemes[i];
                }
                Thread.Sleep(1000);
            }
            return null;
        }
        public static AutomationElement GetBasicTreeItem(AutomationElement aeOutsideDetail)
        {
            //找到树视图
            AutomationElement aetree = aeOutsideDetail.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ClassNameProperty, "Tree"));
            if (null == aetree)
            {
                Console.WriteLine("aetree get fail");
                Thread.Sleep(1000);
                return null;
            }
            //找到ProgressBar
            AutomationElement aeProgressBar = aetree.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ProgressBar));
            if (null == aeProgressBar)
            {
                Console.WriteLine("aeProgressBar get fail");
                Thread.Sleep(1000);
                return null;
            }
            //找到tree
            AutomationElement aetree2 = aeProgressBar.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree));
            if (null == aetree2)
            {
                Console.WriteLine("aetree2 get fail");
                Thread.Sleep(1000);
                return null;
            }

            //找到treeItem
            AutomationElement aetreeItem = aetree2.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
            if (null == aetreeItem)
            {
                Console.WriteLine("aetreeItem get fail");
                Thread.Sleep(1000);
                return null;
            }
            return aetreeItem;
        }
        public static AutomationElement GetSearchText(AutomationElement aeOutsideDetail)
        {
            //找到文本搜索控件
                AutomationElement aeAutoComplete = aeOutsideDetail.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ClassNameProperty, "AutoComplete"));
                if (null == aeAutoComplete)
                {
                    Console.WriteLine("AutoComplete get fail");
                    Thread.Sleep(1000);
                    return null;
                }

                AutomationElement aeAutoCompleteBox = aeAutoComplete.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.AutomationIdProperty, "AutoCompleteBox"));
                if (null == aeAutoCompleteBox)
                {
                    Console.WriteLine("aeAutoCompleteBox get fail");
                    Thread.Sleep(1000);
                    return null;
                }


                return aeAutoCompleteBox;
        }
        public static bool lookdevice(string DeviceName, AutomationElement aeAutoCompleteBox, AutomationElement aeOutsideDetail)
        {
            //通过ValuePattern激活输入框
            ValuePattern vpTextBox1 = (ValuePattern)aeAutoCompleteBox.GetCurrentPattern(ValuePattern.Pattern);
            vpTextBox1.SetValue(DeviceName);
            Thread.Sleep(2000);

            AutomationElement aeEdit = aeAutoCompleteBox.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
            if (null == aeEdit)
            {
                Console.WriteLine("aeEdit get fail");
                Thread.Sleep(1000);
                return false;
            }
            //通过ValuePattern设置TextBox1的值
            ValuePattern vpTextBox2 = (ValuePattern)aeEdit.GetCurrentPattern(ValuePattern.Pattern);
            vpTextBox2.SetValue(DeviceName);
            Thread.Sleep(2000);
            System.Windows.Forms.SendKeys.SendWait("{ENTER}");

            AutomationElement CityNode = GetCityNodeWhichexpand(aeOutsideDetail);
            if (null == CityNode)
            {
                Console.WriteLine("CityNode get fail");
                return false;
            }

            AutomationElement CameraNode = GetCameraNode(CityNode);
            if (null == CameraNode)
            {
                Console.WriteLine("CameraNode get fail");
                return false;
            }

            //下拉滚动条
            if (CameraNode.Current.BoundingRectangle.Bottom > 700)
            {
                Thread.Sleep(2000);

                Console.WriteLine("Node off screen.");
                //找到树视图
                bool isPatternAvailable = (bool)
                       aeOutsideDetail.GetCurrentPropertyValue(AutomationElement.IsScrollPatternAvailableProperty);
                Console.WriteLine(isPatternAvailable);
                AutomationElement aetree = aeOutsideDetail.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ClassNameProperty, "Tree"));
                if (null == aetree)
                {
                    Console.WriteLine("aetree get fail");
                    Thread.Sleep(1000);
                    return true;
                }
                isPatternAvailable = (bool)
                       aetree.GetCurrentPropertyValue(AutomationElement.IsScrollPatternAvailableProperty);
                Console.WriteLine(isPatternAvailable);
                //找到ProgressBar
                AutomationElement aeProgressBar = aetree.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ProgressBar));
                if (null == aeProgressBar)
                {
                    Console.WriteLine("aeProgressBar get fail");
                    Thread.Sleep(1000);
                    return true;
                }
                AutomationElement Treeview = aeProgressBar.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree));
                if (null == aeProgressBar)
                {
                    Console.WriteLine("aeProgressBar get fail");
                    Thread.Sleep(1000);
                    return true;
                }
                isPatternAvailable = (bool)
                       Treeview.GetCurrentPropertyValue(AutomationElement.IsScrollPatternAvailableProperty);
                Console.WriteLine(isPatternAvailable);

                ScrollPattern vpScroll = (ScrollPattern)aetree.GetCurrentPattern(ScrollPattern.Pattern);
                
                Thread.Sleep(2000);
                if (vpScroll.Current.VerticallyScrollable)
                {
                    vpScroll.ScrollVertical(ScrollAmount.LargeIncrement);

                }
            }
            //System.Windows.Point clickablePoint;
            SetCursorPos((int)CameraNode.Current.BoundingRectangle.X, (int)CameraNode.Current.BoundingRectangle.Y);
            

            PONITAPI p = new PONITAPI();
            GetCursorPos(ref p);
            Console.WriteLine("鼠标现在的位置X:{0}, Y:{1}", p.x, p.y);
            Console.WriteLine("Sleep 1 sec...");
            Thread.Sleep(1000);


            Console.WriteLine("在X:{0}, Y:{1} 按下鼠标左键", p.x, p.y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, p.x, p.y, 0, 0);
            Thread.Sleep(10);
            mouse_event(MOUSEEVENTF_LEFTDOWN, p.x, p.y, 0, 0);
            Thread.Sleep(1000);

            Console.WriteLine("在X:{0}, Y:{1} 释放鼠标左键", p.x, p.y);
            mouse_event(MOUSEEVENTF_LEFTUP, p.x, p.y, 0, 0);
            Console.WriteLine("程序结束，按任意键退出....");
            Console.ReadKey();
            
            Thread.Sleep(1000);
            return true;
        }
        public static AutomationElement GetCameraNode(AutomationElement CityNode)
        {
            Console.WriteLine(CityNode.Current.Name,"\r\n");
            AutomationElementCollection aeTreeItemNodes = CityNode.FindAll(TreeScope.Children,
              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
            bool  nextExpandedTreeItem = false;
            AutomationElement NextTreeItem = null;
            for (int i = 0; i < aeTreeItemNodes.Count; i++)
            { 
                 ExpandCollapsePattern ExpandPattern1 = (ExpandCollapsePattern)aeTreeItemNodes[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                 if (ExpandPattern1.Current.ExpandCollapseState == ExpandCollapseState.Expanded)
                 {
                     nextExpandedTreeItem = true;
                     NextTreeItem = aeTreeItemNodes[i];
                     break;
                 }
            }

            if (nextExpandedTreeItem == true)
            {
                return GetCameraNode(NextTreeItem);
            }

            AutomationElement aeCameraNode = CityNode.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ClassNameProperty, "TextBlock"));
            if (null == aeCameraNode)
            {
                Console.WriteLine("aeCameraNode get fail");
                return null;
            }

            Console.WriteLine(aeCameraNode.Current.Name);
            return aeCameraNode;
            

        }
        //created at 2019-10-30
        public static unsafe int OpenCameraNode(AutomationElement aetreeItem, DeviceInfo[] deviceInfo, AutomationElement aeOutsideDetail)
        {    

            //定义设备个数
            int DealDeviceNum = 0;
            //获取所有地市级treeItem
            AutomationElementCollection aeCitytreeItemes = aetreeItem.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
            if (0 == aeCitytreeItemes.Count)
            {
                Console.WriteLine("aeCitytreeItem get 0.");
                Thread.Sleep(1000);
                return 0;
            }
            
            int CityNumber = aeCitytreeItemes.Count;
            //Console.WriteLine(CityNumber);
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
                //currentPattern.Collapse();
                ExpandPattern1.Expand();

                //区级节点操作
                Thread.Sleep(1000);
                //获取该地市区级公司
                AutomationElementCollection aedistrictItemes = aeCitytreeItemes[i].FindAll(TreeScope.Children,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                if (0 == aedistrictItemes.Count)
                {
                    Console.WriteLine("aedistrictItemes get 0.");
                    continue;
                }
                int districtNumber = aedistrictItemes.Count;
                //Console.WriteLine(districtNumber);
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
                    ExpandPattern2.Expand();

                    //交流电级节点操作
                    Thread.Sleep(1000);
                    AutomationElementCollection aeVoltageItemes = aedistrictItemes[j].FindAll(TreeScope.Children,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                    if (0 == aeVoltageItemes.Count)
                    {
                        Console.WriteLine("aedistrictItemes get 0.");
                        continue;
                    }
                    int VoltageNumber = aeVoltageItemes.Count;
                    //Console.WriteLine(VoltageNumber);
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
                        Thread.Sleep(1000);
                        AutomationElementCollection aelineItemes = aeVoltageItemes[k].FindAll(TreeScope.Children,
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                        if (0 == aelineItemes.Count)
                        {
                            Console.WriteLine("aelineItemes get 0.");
                            continue;
                        }
                        int lineNumber = aelineItemes.Count;
                        //Console.WriteLine(lineNumber);
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
                            Thread.Sleep(1000);
                            AutomationElementCollection aetowerItemes = aelineItemes[l].FindAll(TreeScope.Children,
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                            if (0 == aetowerItemes.Count)
                            {
                                Console.WriteLine("aetowerItemes get 0.");
                                continue;
                            }
                            int towerNumber = aetowerItemes.Count;
                            //Console.WriteLine(towerNumber);
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
                                ExpandPattern5.Expand();
                                Thread.Sleep(1000);

                                //摄像头操作
                                AutomationElementCollection aeCameraNode = aetowerItemes[m].FindAll(TreeScope.Children,
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                                if (0 == aeCameraNode.Count)
                                {
                                    Console.WriteLine("aeCameraNode get fail");
                                    continue;
                                }


                                //点击摄像头
                                for (int n = 0; n < aeCameraNode.Count; n++)
                                {
                                     AutomationElement aeCameraButton = aeCameraNode[n].FindFirst(TreeScope.Children,
                                new PropertyCondition(AutomationElement.ClassNameProperty, "TextBlock"));
                                    Console.WriteLine(aeCameraButton.Current.Name);
                                    int o = aeCameraButton.Current.Name.IndexOf('(');
                                    int p = aeCameraButton.Current.Name.IndexOf(')');
                                    int q;
                                    string DName1 = aeCameraButton.Current.Name.Substring(o+1, p-o-1);
                                    Console.WriteLine(DName1);
                                    deviceInfo[DealDeviceNum].DeviceName = DName1;

                                    clickCameraNode(aeCameraButton);
                                    Thread.Sleep(30000);

                                    bool DeviceButtionflag = GetDeviceVideoButton(aeOutsideDetail);
                                    
                                    //等待15秒
                                    Thread.Sleep(1000);
                                    //摄像头视频数据获取
                                    if(DeviceButtionflag)
                                    {
                                        deviceInfo[DealDeviceNum].DeviceState = true;
                                    }
                                    else
                                    {
                                        deviceInfo[DealDeviceNum].DeviceState = false;
                                    }
                                    
                                    DealDeviceNum++;
                                    if (n == aeCameraNode.Count - 1)
                                        ExpandPattern5.Collapse();
                                }
                                 

                                //折叠节点
                                if (m == towerNumber - 1)
                                {
                                    ExpandPattern4.Collapse();

                                }
                            }
                                //折叠节点
                                if (l == lineNumber - 1)
                                {
                                    ExpandPattern3.Collapse();
                                }
                            }
                            //折叠节点
                            if (k == VoltageNumber - 1)
                            {
                                ExpandPattern2.Collapse();
                            }
                        }
                        //折叠节点
                        if (j == districtNumber - 1)
                        {
                            ExpandPattern1.Collapse();
                        }
                    
                }
            }
            return DealDeviceNum;
        }
        //created at 2019-10-30
        public static void clickCameraNode(AutomationElement CameraNode)
        {
            //System.Windows.Point clickablePoint;
            SetCursorPos(((int)CameraNode.Current.BoundingRectangle.Left + (int)CameraNode.Current.BoundingRectangle.Right)/2,
                ((int)CameraNode.Current.BoundingRectangle.Top + (int)CameraNode.Current.BoundingRectangle.Bottom)/2);
            

            PONITAPI p = new PONITAPI();
            GetCursorPos(ref p);
            //Console.WriteLine("鼠标现在的位置X:{0}, Y:{1}", p.x, p.y);
            //Console.WriteLine("Sleep 1 sec...");
            Thread.Sleep(100);


            //Console.WriteLine("在X:{0}, Y:{1} 按下鼠标左键", p.x, p.y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, p.x, p.y, 0, 0);
            Thread.Sleep(200);
            mouse_event(MOUSEEVENTF_LEFTUP, p.x, p.y, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTDOWN, p.x, p.y, 0, 0);
            Thread.Sleep(1000);
            //Console.WriteLine("在X:{0}, Y:{1} 释放鼠标左键", p.x, p.y);
            mouse_event(MOUSEEVENTF_LEFTUP, p.x, p.y, 0, 0);
            //Console.WriteLine("程序结束，按任意键退出....");
            //Console.ReadKey();

        }
        //单击
        public static void singleclickNode(AutomationElement CameraNode)
        {
            //System.Windows.Point clickablePoint;
            SetCursorPos(((int)CameraNode.Current.BoundingRectangle.Left + (int)CameraNode.Current.BoundingRectangle.Right)/2,
                ((int)CameraNode.Current.BoundingRectangle.Top + (int)CameraNode.Current.BoundingRectangle.Bottom)/2);
            

            PONITAPI p = new PONITAPI();
            GetCursorPos(ref p);
            //Console.WriteLine("鼠标现在的位置X:{0}, Y:{1}", p.x, p.y);
            //Console.WriteLine("Sleep 1 sec...");
            Thread.Sleep(100);

            //Console.WriteLine("在X:{0}, Y:{1} 按下鼠标左键", p.x, p.y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, p.x, p.y, 0, 0);
            Thread.Sleep(1000);
            //Console.WriteLine("在X:{0}, Y:{1} 释放鼠标左键", p.x, p.y);
            mouse_event(MOUSEEVENTF_LEFTUP, p.x, p.y, 0, 0);
            //Console.WriteLine("程序结束，按任意键退出....");
            //Console.ReadKey();

        }
        //created at 2019-10-30
        public static AutomationElement GetOnlineVideoTab(AutomationElement aeOutsideDetail)
        { AutomationElement aeTabControl = aeOutsideDetail.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, "DeviceTab"));
            if (null == aeTabControl)
            {
                Console.WriteLine("aeTabControl get fail at GetDeviceInfoRealtime");
                Thread.Sleep(1000);
                return null;
            }
            AutomationElement aeOnlineVideo = aeTabControl.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, "OnlineVideo"));
            if (null == aeOnlineVideo)
            {
                Console.WriteLine("aeOnlineVideo get fail at GetDeviceInfoRealtime");
                Thread.Sleep(1000);
                return null;
            }

            AutomationElement aeOnlineVideoTab = aeOnlineVideo.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, "OnlineVideoTab"));
            if (null == aeOnlineVideoTab)
            {
                Console.WriteLine("aeOnlineVideoTab get fail at GetDeviceInfoRealtime");
                Thread.Sleep(1000);
                return null;
            }
            return aeOnlineVideoTab;
        }
        public static string GetDeviceInfoRealtime(AutomationElement aeOutsideDetail)
        {
            AutomationElement aeOnlineVideoTab = GetOnlineVideoTab(aeOutsideDetail);
            if(aeOnlineVideoTab == null)
            {
                Console.WriteLine("aeOnlineVideoTab get fail at GetDeviceInfoRealtime");
                Thread.Sleep(1000);
                return "";
            }
            AutomationElement aeName = aeOnlineVideoTab.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, "CODE"));
            if (null == aeName)
            {
                Console.WriteLine("aeName get fail at GetDeviceInfoRealtime");
                Thread.Sleep(1000);
                return "";
            }
            return aeName.Current.Name;
        }
        //created at 2019-10-31
        public static bool GetDeviceVideoButton(AutomationElement aeOutsideDetail)
        {
            AutomationElement aeOnlineVideoTab = GetOnlineVideoTab(aeOutsideDetail);
            if(aeOnlineVideoTab == null)
            {
                Console.WriteLine("aeOnlineVideoTab get fail ");
                Thread.Sleep(1000);
                return false;
            }
            AutomationElement aePlaycontrol = aeOnlineVideoTab.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ClassNameProperty, "PlayerControl"));
            if (null == aePlaycontrol)
            {
                Console.WriteLine("aeNaePlaycontrolame get fail ");
                Thread.Sleep(1000);
                return false;
            }

            AutomationElement aePane = aePlaycontrol.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ClassNameProperty, "WindowsFormsHost"));
            if (null == aePane)
            {
                Console.WriteLine("aePane get fail ");
                Thread.Sleep(1000);
                return false;
            }

            AutomationElement aeMedia = aePane.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, "Media"));
            if (null == aeMedia)
            {
                Console.WriteLine("aeMedia get fail ");
                Thread.Sleep(1000);
                return false;
            }

            AutomationElement aeMainPanel = aeMedia.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, "MainPanel"));
            if (null == aeMainPanel)
            {
                Console.WriteLine("aeMainPanel get fail ");
                Thread.Sleep(1000);
                return false;
            }

            AutomationElement aeTooltip = aeMainPanel.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, "TooltipPanel"));
            if (null == aeTooltip)
            {
                Console.WriteLine("aeTooltip get fail ");
                Thread.Sleep(1000);
                return false;
            }

            AutomationElement aePicPanel = aeTooltip.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, "PicPanel"));
            if (null == aePicPanel)
            {
                Console.WriteLine("aePicPanel get fail ");
                Thread.Sleep(1000);
                return false;
            }
            
            AutomationElement aeRecordbutton = aePicPanel.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.AutomationIdProperty, "btn_luxiang"));
            if (null == aeRecordbutton)
            {
                Console.WriteLine("aeRecordbutton get fail ");
                Thread.Sleep(1000);
                return false;
            }

            singleclickNode(aeRecordbutton);
            Thread.Sleep(2000);
            singleclickNode(aeRecordbutton);
            Thread.Sleep(6000);
            //有弹出窗口表示设备在线，没有窗口表示不在线
            return IfhaveSaveaswindow();
        }
        public static bool GetDevicestate(string DeviceName, DeviceInfo[] deviceInfo, int DeviceNum)
        {
            for(int i = 0; i < DeviceNum; i++)
            {
                if(deviceInfo[i].DeviceName == DeviceName)
                {
                    return deviceInfo[i].DeviceState;
                }
            }
            return false;
            
        }
        //created at 2019-10-31
        public static bool IfhaveSaveaswindow()
        {
            //根据执行程序名获取进程
                Process[] p2 = new Process[2];
                p2 = Process.GetProcessesByName("IoTPlatform");
                if (null == p2[0])
                {
                    Console.WriteLine("Process get fail");
                    return false;
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
                    return false;
                }
                Console.WriteLine("Process 2");

                AutomationElement aewindow = aeForm.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.NameProperty, "另存为"));
                if (null == aewindow)
                {
                    Console.WriteLine("aewindow get fail ");
                    Thread.Sleep(1000);
                    return false;
                }
                //获得对主窗体对象的引用
                AutomationElement aeline = aewindow.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TitleBar));
                if (null == aeline)
                {
                    Console.WriteLine("aeline get fail ");
                    Thread.Sleep(1000);
                    return false;
                }

                AutomationElement aebutton = aeline.FindFirst(TreeScope.Children,
              new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
                if (null == aebutton)
                {
                    Console.WriteLine("aebutton get fail ");
                    Thread.Sleep(1000);
                    return false;
                }

                singleclickNode(aebutton);
                return true;

        }
    }
}