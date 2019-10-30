#define IOT
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
        }

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
                //DeviceName[0] = "100210001403510098";
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

                for (int Index = 0; Index < DeviceNUM; Index++)
                {
                    //搜索设备，查看视频状态, 记录状态
                    if (false == lookdevice(DeviceName[Index], aeSearch,aeOutsideDetail))
                    {
                        DeviceState[Index] = false;
                        Console.WriteLine("Device Index",Index,"can't find");

                    }
                    else
                    {
                        DeviceState[Index] = true;//暂时这样
                        Console.WriteLine("Device Index", Index, "can find");
                        //获取到摄像头节点
                        //折叠主节点
                        ExpandCollapsePattern ExpandPattern1 = (ExpandCollapsePattern)aeBasicTreeItem.GetCurrentPattern(ExpandCollapsePattern.Pattern);

                        ExpandPattern1.Collapse();
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
    }
}