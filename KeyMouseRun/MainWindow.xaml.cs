using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;

namespace KeyMouseRun
{
    /// <summary>
    /// 存储数据的节点。
    /// </summary>
    public struct RunNode
    {
        //NO 原为string
        public byte NO;
        public uint TIME;
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// 字典路径。可直接移动到方法内。
        /// </summary>
        private string jsonDicFilePath = System.IO.Directory.GetCurrentDirectory();
        /// <summary>
        /// Json数据读取。可直接移动到方法内。
        /// </summary>
        private JObject jsonDBInfo;

        /// <summary>
        /// 执行内容
        /// </summary>
        private readonly List<RunNode> runRes = new List<RunNode>();
        /// <summary>
        /// 字典
        /// </summary>
        private readonly Dictionary<string, byte> keyMouseDic = new Dictionary<string, byte>();
        /// <summary>
        /// 鼠标坐标的X值
        /// </summary>
        private int mouseX = 0;
        /// <summary>
        /// 鼠标坐标的Y值
        /// </summary>
        private int mouseY = 0;
        /// <summary>
        /// 避免重复运行的flag
        /// </summary>
        private bool RunBusy = false;

        /// <summary>
        /// 鼠标事件。
        /// </summary>
        [DllImport("user32.dll")]
        private static extern void mouse_event(byte dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
        /// <summary>
        /// 键盘事件。
        /// </summary>
        [DllImport("user32.dll", EntryPoint = "keybd_event", SetLastError = true)]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);
        /// <summary>
        /// 设置鼠标位置
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        private static extern void SetCursorPos(int x, int y);
        /// <summary>
        /// 
        /// </summary>
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool SetProcessWorkingSetSize(IntPtr proc, int min, int max);

        /// <summary>
        /// 控件左上角
        /// </summary>
        private System.Windows.Point pt;

        public MainWindow()
        {
            InitializeComponent();

            DicInit();

            Thread threadGC = new Thread(FlushMemory)
            {
                IsBackground = true
            };
            threadGC.Start();

        }

        /// <summary>
        /// 清理内存的方法。
        /// </summary>
        private void FlushMemory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
            }
            Thread.Sleep(5000);
            FlushMemory();
        }

        /// <summary>
        /// 窗口拖动。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackRect_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
                Thread.Sleep(100);//降低占用
            }
        }

        /// <summary>
        /// 运行程序，始终执行单一程序。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            if (RunBusy == false)
            {
                RunBusy = true;

                Thread runThread = new Thread(KeyMouseRun)
                {
                    IsBackground = false
                };
                runThread.Start();
            }
        }

        /// <summary>
        /// 选择字典。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectBtn_Click(object sender, RoutedEventArgs e)
        {
            DataInit();
        }
        
        /// <summary>
        /// 主要操作。
        /// </summary>
        private void KeyMouseRun()
        {
            if (runRes.Count == 0)
            {
                MessageBox.Show("Null");
                MouseMove2SelectMidPosition();
                RunBusy = false;
                return;
            }
            try
            {
                foreach (var item in runRes)
                {
                    switch (item.NO)
                    {
                        case 0:
                            Thread.Sleep((int)item.TIME * 100);
                            break;
                        case byte.MaxValue - 20:
                            Dispatcher.Invoke(new Action(() => {
                                WindowState = WindowState.Minimized;
                            }));
                            break;
                        case byte.MaxValue - 21:
                            Dispatcher.Invoke(new Action(() => {
                                System.Windows.Application.Current.Shutdown();
                            }));
                            return;
                        case byte.MaxValue - 2:
                            mouse_event(keyMouseDic["MOUSEEVENTF_LEFTDOWN"], 0, 0, 0, 0);
                            break;
                        case byte.MaxValue - 4:
                            mouse_event(keyMouseDic["MOUSEEVENTF_LEFTUP"], 0, 0, 0, 0);
                            break;
                        case byte.MaxValue - 8:
                            mouse_event(keyMouseDic["MOUSEEVENTF_RIGHTDOWN"], 0, 0, 0, 0);
                            break;
                        case byte.MaxValue - 16:
                            mouse_event(keyMouseDic["MOUSEEVENTF_RIGHTUP"], 0, 0, 0, 0);
                            break;
                        case byte.MaxValue - 17:
                            mouseX = (int)item.TIME;
                            break;
                        case byte.MaxValue - 18:
                            mouseY = (int)item.TIME;
                            break;
                        case byte.MaxValue - 19:
                            SetCursorPos(mouseX, mouseY);
                            break;
                        default:
                            if (item.TIME == 0 || item.TIME == 2)
                                keybd_event(item.NO, 0, item.TIME, 0);
                            //tmpStr += item.NO.ToString() + " " + item.TIME.ToString() + "\n";
                            break;
                    }
                }

                MouseMove2RunMidPosition();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误");
            }
            finally
            {
                RunBusy = false;
            }
        }

        /// <summary>
        /// 字典初始化。
        /// </summary>
        private void DicInit()
        {
            try
            {
                //设置文件路径
                string dicPathStr = System.Environment.CurrentDirectory.ToString();
                dicPathStr = dicPathStr.Substring(0, dicPathStr.LastIndexOf(@"\"));
                dicPathStr = dicPathStr.Substring(0, dicPathStr.LastIndexOf(@"\"));
                dicPathStr = dicPathStr.Substring(0, dicPathStr.LastIndexOf(@"\"));
                //获取最终路径
                jsonDicFilePath = dicPathStr + @"\Src\KeyMouseDic.json";


                //数据初始化
                using System.IO.StreamReader file = System.IO.File.OpenText(jsonDicFilePath);
                using JsonTextReader reader = new JsonTextReader(file);
                jsonDBInfo = (JObject)JToken.ReadFrom(reader);
                // i 分类
                foreach (var i in jsonDBInfo)
                {
                    keyMouseDic.Add(i.Key, (byte)i.Value);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                System.Windows.Application.Current.Shutdown();
            }
        }

        /// <summary>
        /// 操作初始化/更新。
        /// </summary>
        private void DataInit()
        {
            try
            {
                //清空原数据
                runRes.Clear();
                string dataPath = GetFilePath();
                //获取失败
                if (string.IsNullOrEmpty(dataPath))
                {
                    //鼠标移到选择按钮 重新选择运行数据资源
                    MouseMove2SelectMidPosition();
                    return;
                }
                //记录需要运行的内容
                XDocument xmldoc = XDocument.Load(dataPath);
                XElement xmlele = xmldoc.Root;
                IEnumerable<XElement> tmpele = xmlele.Elements();
                foreach (var item in tmpele)
                {
                    runRes.Add(new RunNode
                    {
                        //M开头倒记
                        NO = item.Name.ToString()[0] != 'M' ? keyMouseDic[item.Name.ToString()] : (byte)(byte.MaxValue - keyMouseDic[item.Name.ToString()]),
                        TIME = uint.Parse(item.Value)
                    });
                }
                //鼠标移到运行按钮
                MouseMove2RunMidPosition();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                System.Windows.Application.Current.Shutdown();
            }
        }

        /// <summary>
        /// 选择文件。
        /// </summary>
        /// <returns></returns>
        private string GetFilePath()
        {
            try
            {
                //获取基本路径
                string exePath = System.Environment.CurrentDirectory.ToString();
                exePath = exePath.Substring(0, exePath.LastIndexOf(@"\"));
                exePath = exePath.Substring(0, exePath.LastIndexOf(@"\"));
                exePath = exePath.Substring(0, exePath.LastIndexOf(@"\"));

                Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
                openFileDialog.Title = "选择工具运行字典";
                openFileDialog.Filter = "xml字典|*.xml";
                openFileDialog.InitialDirectory = exePath + @"\Src";
                openFileDialog.FileName = string.Empty;
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = false;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.DefaultExt = "xml";
                if (openFileDialog.ShowDialog() == true)
                {
                    return openFileDialog.FileName;
                }
                //MessageBox.Show(openFileDialog.FileName.Substring(openFileDialog.FileName.LastIndexOf(@".")));
                
            }
            catch (Exception)
            {
                MessageBox.Show("数据获取失败.");
            }

            return "";
        }

        /// <summary>
        /// 鼠标移到运行的位置。
        /// </summary>
        private void MouseMove2RunMidPosition()
        {
			//最小化 无控件位置和控件Width Height 待优化
            Dispatcher.Invoke(new Action(() => {
                //do something here
                pt = RunBtn.PointToScreen(new System.Windows.Point(0, 0));

                if (pt.X > 0 && pt.Y > 0
                    && pt.X <= SystemParameters.PrimaryScreenWidth
                    && pt.Y <= SystemParameters.PrimaryScreenHeight)
                {
                    SetCursorPos((int)pt.X + (int)RunBtn.Width / 2, (int)pt.Y + (int)RunBtn.Height / 2);
                }
            }));
        }

        /// <summary>
        /// 鼠标移到选择的位置。
        /// </summary>
        private void MouseMove2SelectMidPosition()
        {
            Dispatcher.Invoke(new Action(() => {
                //do something here
                pt = SelectBtn.PointToScreen(new System.Windows.Point(0, 0));

                if (pt.X > 0 && pt.Y > 0
                    && pt.X <= SystemParameters.PrimaryScreenWidth
                    && pt.Y <= SystemParameters.PrimaryScreenHeight)
                {
                    SetCursorPos((int)pt.X + (int)SelectBtn.Width / 2, (int)pt.Y + (int)SelectBtn.Height / 2);
                }
            }));
        }
    }
}
