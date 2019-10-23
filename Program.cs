using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation.Provider;
using System.Windows.Automation.Text;
using System.Windows.Automation;


namespace UIAutomationTest
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("\nBegin WinForm UIAutomation test run\n");
                // launch Form1 application
                // get refernce to main Form control
                // get references to user controls
                // manipulate application
                // check resulting state and determine pass/fail

                Console.WriteLine("\nBegin WinForm UIAutomation test run\n");
                Console.WriteLine("Launching WinFormTest application");

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

                //修改登录界面 用户名、密码、环境信息； 点击确认键连接服务器
                                Thread.Sleep(10000);
                AutomationElement aeForm = AutomationElement.FromHandle(p2[0].MainWindowHandle);
                //获得对主窗体对象的引用，该对象实际上就是 Form1 应用程序(方法一)
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

                //根据controltype获取到服务平台
                AutomationElement aeForm = aeDeskTop.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.NameProperty, "输变电物联网服务平台"));
                if (null == aeForm)
                {
                    Console.WriteLine("aeForm get fail");
                }
                
                //获取子窗口
                AutomationElement aeTabControl = aeForm.FindFirst(TreeScope.Children,
                 new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab));
                if (null == aeTabControl)
                {
                    Console.WriteLine("aeTabControl get fail");
                }

                //获取输电可视化窗口
                AutomationElement aeTabItemControl = aeTabControl.FindFirst(TreeScope.Children,
                 new PropertyCondition(AutomationElement.NameProperty, "输电线路可视化"));
                if (null == aeTabItemControl)
                {
                    Console.WriteLine("aeTabItemControl get fail");
                }

                //获取自定义窗口 
                AutomationElement aeCustomControl = aeTabItemControl.FindFirst(TreeScope.Children,
                 new PropertyCondition(AutomationElement.ClassNameProperty, "OutsidePage"));
                if (null == aeCustomControl)
                {
                    Console.WriteLine("aeCustomControl get fail");
                }
                
                //找到详情的Button控件
                AutomationElement aeButton = aeCustomControl.FindFirst(TreeScope.Children,
                  new PropertyCondition(AutomationElement.HelpTextProperty, "详情"));
                if (null == aeButton)
                {
                    Console.WriteLine("aeButton get fail");
                }
                //通过InvokePattern模拟点击按钮
                InvokePattern ipClickButton = (InvokePattern)aeButton.GetCurrentPattern(InvokePattern.Pattern);
                ipClickButton.Invoke();

                Thread.Sleep(1500);

                {
                    Console.WriteLine("Did not find it.");
                    Console.WriteLine("Test scenario: *FAIL*");
                }

                Console.WriteLine("wait for long time.");
                Thread.Sleep(100000);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex.Message);
            }
        }
    }
}