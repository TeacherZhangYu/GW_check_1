using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.Structure;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System.Threading;
using System.Numerics;
using System.Threading;
using GW_check;
using Yolov5Net.Scorer;
using Yolov5Net.Scorer.Models;
using MSWord = Microsoft.Office.Interop.Word;
using MSExcel= Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace KJ_chenk
{
    public partial class MainForm : Form
    {
        #region  所有全局变量声明     
        public static MainForm frm = new MainForm();
        private DataBase database = new DataBase();          //数据库连接
        UserLoad userform = new UserLoad();                  //用户登陆界面
        private static UserControlVideoPlayer videoPlayer;   //视频播放控件
        private VideoCapture capture;
        private static imagePlayer imageInteract;
        imagePlayer imageplayer = new imagePlayer();
        public int outCheckHeight;
        public int outCheckWidth;


        //各个文件需要保存的路径
        public string xmlRoute = System.Environment.CurrentDirectory + "\\xmlFlie\\userdata.xml";  //保存xml文件的路径，用的当前debug文件夹下的目录
        public string image_wait_route = System.Environment.CurrentDirectory + "\\imageWait\\";      
        public string image_handle_route = System.Environment.CurrentDirectory + "\\imageHandle\\";
        public string image_waite_handle = System.Environment.CurrentDirectory + "\\waitToHandle\\";
        public string onnx_file = System.Environment.CurrentDirectory + "\\onnx\\";                     //onnx所在的路径可以动态修改
        string str_path = System.Environment.CurrentDirectory + "\\modelFile\\model.docx";
        //生成word文档路径和对象
        object wordPath;
        public string report_file = System.Environment.CurrentDirectory + "\\reportFile\\";         //该位置可以用于存放报告的word文档以及excel文件
        
        MSWord.Document wordDoc;                       //Word文档变量                                                                           
        Object Nothing = Missing.Value;                //由于使用的是COM库，因此有许多变量需要用Missing.Value代替
        object unite = MSWord.WdUnits.wdStory;         //写入黑体文本

        //生成excel文档路径和变量
        MSWord.Application word_app = new MSWord.Application();
        MSExcel.Application excel_app = new MSExcel.Application();

        public UInt32 excel_count = 2;                //计cxcel表中的行数，从第二行开始填写
        public UInt32 image_handle_count = 0;          //记录处理过的图像的数量
        private static string wordName;               //excel表格的文件名
        private string batchFileName;                 //批量检测的时候，检测excel文件的命名

        //public string video_wait_route = "E:\\公司项目\\检测文件\\";
        private Xml xmlfile = new Xml();                     //实例化xml文件类的对象
        private string videoFile;
        public static string fileNameWithoutExtension;        //获取视频的文件名
        public UInt32 Python_Count = 0;                       //截图命名的起始值
        public UInt32 PythonCountHandle = 0;
        public uint imageCycle = 50;                          //设置为多少帧进行截图
        public int imageShow = 1000;                          //设置当前参数为读取与处理文件夹截图的速度
        List<string> lstr = new List<string>();
        private static int i = 0;                             //检测视频的编号
        public bool endMssege = false;
        static string batchVideoPath=null;
        #endregion
        public MainForm()
        {
            InitializeComponent();
            frm = this;
        }
        public void getSize()
        {
            outCheckHeight = out_check.Height;
            outCheckWidth = out_check.Width;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            userform.ShowDialog();
            if (userform.DialogResult == DialogResult.OK)
            {
                userform.Dispose();
            }
            else if (userform.DialogResult == DialogResult.Cancel)
            {
                this.Close();
            }
            initText();
        }
        private void initText()  ///初始化自动检测界面所有的元器件textbox以及各种panel
        {
            timer_auto_Video.Enabled = false;   //自动播放使能初始化
        }
        #region  数据库连接
        private void btn_link_Click(object sender, EventArgs e)
        {
            //编写数据库连接串
            string connStr = "Data source=.;Initial Catalog=test;User ID=sa;Password=pwdpwd";
            //创建SqlConnection的实例
            database.mysqlconnection(ipName.Text, 3306, userName.Text, databaseName.Text);
        }
        //数据库的数据按日期查询
        private void btn_query_Click(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            var stu = database.QueryData(ipName.Text, 3306, userName.Text, databaseName.Text);
            dataGridView1.DataSource = stu;
        }
        #endregion
        #region  自动检测界面按钮操作
        private void loadVideo_Click(object sender, EventArgs e)
        {
            if (btn_solo.Checked)
            {
                var newVideo = new OpenFileDialog();
                videoPlayer = null;
                newVideo.Title = "选择文件";
                newVideo.Filter = "视频文件|*.avi;*.rmvb;*.rm;*.mp4";
                DeleteFolder(image_handle_route);
                newVideo.RestoreDirectory = true;
                newVideo.Multiselect = false;          //多选关闭
                if (newVideo.ShowDialog() == DialogResult.OK)
                {
                    videoFile = newVideo.FileName;
                    videoPlayer = new UserControlVideoPlayer(pure_video.Size, videoFile);
                    pure_video.Controls.Clear();
                    pure_video.Controls.Add(videoPlayer);
                }
                else
                {
                    return;
                }
                if (videoPlayer != null)
                {
                    videoPlayer.newsize(pure_video.Size);
                    readVideoName();
                    videoPlayer.onVideoEnd += VideoEndChangeButton;
                }
            }
            else if (btn_batch.Checked)
            {

                lstr.Clear();
                i = 0;

                //videoPlayer = null;                        //每一次导入视频都要播放器初始化
                DeleteFolder(image_handle_route);            //载入视频前进行文件夹格式化
                FolderBrowserDialog dialog = new FolderBrowserDialog();   //选择一个文件夹
                dialog.Description = "请选择图片所在文件夹";              
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    batchVideoPath = dialog.SelectedPath;
                }
                string s;
                string realPath = batchVideoPath + "\\";
                DirectoryInfo d = new DirectoryInfo(realPath);
                FileInfo[] files = d.GetFiles("*.mp4");      //可以选择文件的格式

                foreach (FileInfo file in files)       //foreach遍历文件夹
                {
                    s = file.FullName;
                    lstr.Add(s);                      //将文件夹内的XML文件的路径依次存入链表中
                }
            }
            else
            {
                MessageBox.Show("请选择检测模式！！！");
            }
        }
        private void VideoEndChangeButton()
        {
            //videoStart.Enabled = true;
            loadVideo.Enabled = false;
            videoStart.Text = "开始检测";
        }
        private void readVideoName()
        {

            fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(videoFile);
            string[] strArray = fileNameWithoutExtension.Split(new char[] { '-' });  //利用函数分割两端的视频名称填写在textbox中
            startPiont.Text = strArray[0];
            endPoint.Text = strArray[1];
            //for (int i = 0; i < fileNameWithoutExtension.Length;i++)  //读取视频的名称与上述的方式不同，如有需求直接填写
            //{
            //    char str = fileNameWithoutExtension[i];
            //    //str.ToString();
            //    //int result= String.Compare(str, '-');
            //    bool result = str.Equals('-');
            //    if (result)
            //    {
            //      startPiont.Text = fileNameWithoutExtension.Substring(0, i);
            //      endPoint.Text = fileNameWithoutExtension.Substring(i+1, i);                    
            //    }
            //}
        }
        private void videoStart_Click(object sender, EventArgs e)
        {
            // 启动线程       
            if (btn_solo.Checked)
            {
               
                if (videoPlayer == null)
                {
                    MessageBox.Show("请先载入视频！！！");
                    return;
                }
                if (videoStart.Text == "开始检测")
                {
                    if (File.Exists(report_file + fileNameWithoutExtension + ".xlsx"))            //这里注意后期要改
                    {
                        MessageBox.Show("该视频已经检测，请勿重复检测！！！");
                        return;
                    }
                    //给输出图片的定时器使能
                    timerTovideo.Enabled = true;
                    //buttonDataRetrive_Click(sender, e);
                    loadVideo.Enabled = false;
                    creat_excel.Enabled = false;
                    generate_report.Enabled = false;
                    videoStart.Text = "暂停检测";
                    pure_video.Controls.Clear();
                    pure_video.Controls.Add(videoPlayer);
                    timer_auto_Video.Enabled = true;        //相关的定时器使能，包括截图、添加excel表格参数、抽帧
                    timerToImage.Enabled = true;
                    PythonCountHandle = 0;
                    image_handle_count = 0;
                    timerToExcel.Enabled = true;
                    timerToImage.Interval = imageShow;
                    excel_count = 2;
                    AutoEcxel();
                    videoPlayer.startVideo();
                }
                else
                {
                    //videoStart.Enabled = true;
                    loadVideo.Enabled = true;
                    creat_excel.Enabled = true;
                    generate_report.Enabled = true;
                    videoStart.Text = "开始检测";
                    videoPlayer.pauseVideo();
                    timer_auto_Video.Enabled = false;                 
                    Python_Count = 0;
                    Thread.Sleep(2500);               //当前线程由于运行时会占用检测的线程，所以要延迟2秒防止线程阻塞
                    timerToImage.Enabled = false;
                    timerToExcel.Enabled = false;
                    Thread.Sleep(1000);
                    DeleteFolder(image_wait_route);
                }
            }
            else if (btn_batch.Checked)
            {
                var newVideo = new OpenFileDialog();
                if (videoStart.Text == "开始检测")
                {
                    try
                    {
                        excel_count = 2;
                        AutoEcxel();
                        timerToExcel.Enabled = true;
                        loadVideo.Enabled = false;
                        creat_excel.Enabled = false;
                        generate_report.Enabled = false;
                        videoStart.Text = "暂停检测";
                        AutoVideo();                       
                        timer_auto_Video.Enabled = true;
                        timerToImage.Enabled = true;
                        PythonCountHandle = 0;
                        image_handle_count = 0;
                        timerToExcel.Enabled = true;
                        timerToImage.Interval = imageShow;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
                else
                {
                    loadVideo.Enabled = true;
                    creat_excel.Enabled = true;
                    generate_report.Enabled = true;
                    videoStart.Text = "开始检测";
                    videoPlayer.pauseVideo();                    
                    Thread.Sleep(2500);                   
                    timer_auto_Video.Enabled = false;
                    timerToImage.Enabled = false;
                    timerToExcel.Enabled = false;
                    timerToExcel.Enabled = false;
                    Thread.Sleep(1000);
                    DeleteFolder(image_wait_route);
                }
            }
            else
            {
                MessageBox.Show("请先选择检测模式！！！");
            }           
        }         
        private void DeleteFolder(string directory)              //删除指定路径的文件
        {
            foreach (string route in Directory.GetFileSystemEntries(directory))
            {
                if (File.Exists(route))
                {
                    try
                    {
                        FileInfo fi = new FileInfo(route);
                        if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                        fi.Attributes = FileAttributes.Normal;
                        File.Delete(route);//直接删除其中的文件 
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.ToString());
                    }
                }
            }
        }
        private void creat_excel_Click(object sender, EventArgs e)    //生成excel文件只需要
        {
            MessageBox.Show("表格已成功生成！！！");    //这里的导出表格只需要将表格转移存储的位置，然后有一个弹窗提示
        }
        public void AutoEcxel()                                     //自动生成excel文件，开始检测按钮点击的时候就要生成
        {
            string excelFileName;
            //保存excel文档数据           
            //创建excel文档
            MSExcel.Application excel_app = new MSExcel.Application();
            MSExcel.Workbook excel_book = excel_app.Workbooks.Add();
            MSExcel.Worksheet excel_sheet = (MSExcel.Worksheet)excel_book.ActiveSheet;
            batchFileName = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
            if (btn_batch.Checked)
            {
                excel_sheet.Name = System.IO.Path.GetFileNameWithoutExtension(batchFileName);
            }
            else
            {               
                excel_sheet.Name = System.IO.Path.GetFileNameWithoutExtension(videoFile);    //
            }           
            //在第一行填写标题
            excel_sheet.Cells[1, 1] = "工程编码";
            excel_sheet.Cells[1, 2] = "起点编码";
            excel_sheet.Cells[1, 3] = "终点编码";
            excel_sheet.Cells[1, 4] = "检测对象";
            excel_sheet.Cells[1, 5] = "管段直径";
            excel_sheet.Cells[1, 6] = "官网材质";
            excel_sheet.Cells[1, 7] = "检测图像";
            excel_sheet.Cells[1, 8] = "检测视频";
            excel_sheet.Cells[1, 9] = "缺陷类型";
            excel_sheet.Cells[1, 10] = "缺陷名称";
            excel_sheet.Cells[1, 11] = "缺陷位置";
            excel_sheet.Cells[1, 12] = "缺陷级别";
            excel_sheet.Cells[1, 13] = "检测地点";
            excel_sheet.Cells[1, 14] = "检测单位";
            excel_sheet.Cells[1, 15] = "检测人员";
            excel_sheet.Cells[1, 16] = "检测日期";
            excel_sheet.Cells[1, 17] = "检测长度";
            excel_sheet.Cells[1, 18] = "路面类型";
            excel_sheet.Cells[1, 19] = "检测方式";
            excel_sheet.Cells[1, 20] = "养护建议";
            excel_sheet.Cells[1, 21] = "养护位置";
            excel_sheet.Cells[1, 22] = "检测报告";
            excel_sheet.Cells[1, 23] = "报告位置";
            // 向Worksheet中添加数据
            // 该部分代码就是实现相关部分信息填写

            if (btn_batch.Checked)
            {
                excel_book.SaveAs(report_file + batchFileName + ".xlsx");
            }
            else
            {
                excel_book.SaveAs(report_file + fileNameWithoutExtension + ".xlsx");     //
            }
            //这里目前的想法如果单个视频检测就以视频的名称命名该文件，如果是批量检测就以文件夹的名称来命名
            // 关闭Excel应用程序
            excel_book.Close();
            excel_app.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        private void AddEcxel()
        {
            MSExcel.Application excel_app = new MSExcel.Application();
            string excel_name;
            string excel_file;
            if (btn_batch.Checked)
            {
                excel_name = batchFileName;
                excel_file = batchVideoPath + "\\" + batchFileName;
            }
            else
            {
                excel_name = fileNameWithoutExtension;
                excel_file = videoFile;
            }
            MSExcel.Workbook excel_book = excel_app.Workbooks.Open(report_file + excel_name + ".xlsx");
            MSExcel.Worksheet excel_sheet = (MSExcel.Worksheet)excel_book.ActiveSheet;
            excel_sheet.Name = System.IO.Path.GetFileNameWithoutExtension(excel_file);

            excel_sheet.Cells[excel_count, 1] = projectData.Text;    //工程编码
            excel_sheet.Cells[excel_count, 2] = startPiont.Text;     //起点编码
            excel_sheet.Cells[excel_count, 3] = endPoint.Text;
            excel_sheet.Cells[excel_count, 4] = check_length.Text;
            excel_sheet.Cells[excel_count, 5] = pipe_dia.Text;
            excel_sheet.Cells[excel_count, 6] = pipeMat.Text;
            excel_sheet.Cells[excel_count, 7] = check_image.Text;
            excel_sheet.Cells[excel_count, 8] = excel_sheet.Name;
            excel_sheet.Cells[excel_count, 9] = flaw_type.Text;
            excel_sheet.Cells[excel_count, 10] = "/";
            excel_sheet.Cells[excel_count, 11] = flaw_location.Text;
            excel_sheet.Cells[excel_count, 12] = flaw_level.Text;
            excel_sheet.Cells[excel_count, 13] = check_loc.Text;
            excel_sheet.Cells[excel_count, 14] = checkUnit.Text;
            excel_sheet.Cells[excel_count, 15] = checkName.Text;
            excel_sheet.Cells[excel_count, 16] = dateTime.Value.ToShortTimeString();
            excel_sheet.Cells[excel_count, 17] = check_length.Text;
            excel_sheet.Cells[excel_count, 18] = roadType.Text;
            excel_sheet.Cells[excel_count, 19] = "/";
            excel_sheet.Cells[excel_count, 20] = "/";
            excel_sheet.Cells[excel_count, 21] = "/";
            excel_sheet.Cells[excel_count, 22] = fileNameWithoutExtension;       //报告名称
            excel_sheet.Cells[excel_count, 23] = wordPath;       //报告的额名称以及路径

            //保存文件并且退出应用程序
            excel_book.Save();
            excel_book.Close();
            excel_app.Quit();
        }
        private void timerToExcel_Tick(object sender, EventArgs e)       //自动检测文件中图片数量的变化，填写excel文件中的信息
        {
            
            if (File.Exists(image_handle_route + image_handle_count.ToString() + "_0.jpg"))
            {
                AddEcxel();
                image_handle_count++;
                excel_count++;
            }                     
        }
        #endregion
        #region  xml文件的增删改查用户权限的注册与分配
        private void btn_register_Click(object sender, EventArgs e)
        {
            xmlfile.AppendNode(addUser.Text, userPassword.Text, userWork.Text);
        }
        private void btn_cancel_Click(object sender, EventArgs e)
        {
            addUser.Text = "";
            userPassword.Text = "";
            userWork.Text = "";
        }
        #endregion
        #region   实现视频的自动播放线程        
        private void AutoVideo()
        {          
            if (i == lstr.Count)
            {
                timer_auto_Video.Enabled = false;
                i = 0;         
                Thread.Sleep(2000);
                timerToImage.Enabled = false;
                DeleteFolder(image_wait_route);
                loadVideo.Enabled = true;
                creat_excel.Enabled = true;
                generate_report.Enabled = true;
                MessageBox.Show("所有视频检测完成，可根据需要生成报告！");
            }
            else
            {
                var newVideo = new OpenFileDialog();
                newVideo.FileName = lstr[i];
                videoFile = newVideo.FileName;
                videoPlayer = new UserControlVideoPlayer(pure_video.Size, videoFile);
                pure_video.Controls.Clear();
                pure_video.Controls.Add(videoPlayer);
                if (videoPlayer != null)
                {
                    videoPlayer.newsize(pure_video.Size);
                    readVideoName();
                    videoPlayer.onVideoEnd += VideoEndChangeButton;
                }
                videoPlayer.startVideo();
                loadVideo.Enabled = false;
                videoStart.Text = "暂停检测";
                i++;
            }
        }
        private void timer_auto_Video_Tick(object sender, EventArgs e)         //自动播放文件夹中的视频
        {
            try
            {
                //Mat frame = new Mat();
                if (btn_batch.Checked)
                {
                    if (endMssege)
                    {
                        AutoVideo();
                    }
                    else
                    {
                        // videoStart.Text = "开始检测";
                    }
                }
                else if (btn_solo.Checked)
                {
                    if (endMssege)
                    {
                        timer_auto_Video.Enabled = false;
                        
                        Thread.Sleep(2500);
                        timerToImage.Enabled = false;
                        DeleteFolder(image_wait_route);                 //这里删除视频截图，已经检测过了
                        loadVideo.Enabled = true;
                        creat_excel.Enabled = true;
                        generate_report.Enabled = true;
                        if(Directory.GetDirectories(image_handle_route).Length > 0 || Directory.GetFiles(image_handle_route).Length > 0)
                        {
                            MessageBox.Show("视频检测完成！可根据需要生成报告！");
                        }
                        else
                        {
                            MessageBox.Show("视频检测完成，不存在缺陷无报告生成！！！");
                        }                       
                    }
                }
                else
                {
                    MessageBox.Show("请选择检测模式");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion  
        #region      调用深度学习模型输出结果
        private void timerToImage_Tick(object sender, EventArgs e)      //主要用于检测图片展示图片和信息，填写excel表中的数据
        {
            if (Directory.GetDirectories(image_wait_route).Length > 0 || Directory.GetFiles(image_wait_route).Length > 0)
            {
                getSize();
                if (File.Exists(image_wait_route + Python_Count.ToString() + "_0.jpg"))   //判断该路径下有没有该文件，注意一定要加上文件后缀名
                {                                   
                    string srcPath = image_wait_route + Python_Count.ToString() + "_0.jpg";
                    var image = Image.FromFile(srcPath);
                    var scorer = new YoloScorer<YoloCocoP5Model>(onnx_file + "best20230330.onnx");
                    List<YoloPrediction> predictions = scorer.Predict(image);
                    string flawName = "";
                    int flawNumber = predictions.Count;
                    List<string> list = new List<string>(); 
                    foreach (var prediction in predictions)
                    {
                        string quexianName = prediction.Label.Name;
                        list.Add(quexianName);
                    }
                    for (int i = 0; i < flawNumber; i++)
                    {
                        switch (list[i])
                        {
                            case "TJ":
                                flawName += list[i] + "(脱节)，";
                                break;
                            case "AJ":
                                flawName += list[i] + "(支管暗接)，";
                                break;
                            case "BX":
                                flawName += list[i] + "(变形)，";
                                break;
                            case "CK":
                                flawName += list[i] + "(错口)，";
                                break;
                            case "QF":
                                flawName += list[i] + "(起伏)，";
                                break;
                            case "SL":
                                flawName += list[i] + "(渗漏)，";
                                break;
                            case "FS":
                                flawName += list[i] + "(腐蚀)，";
                                break;
                            case "TL":
                                flawName += list[i] + "(接口材料脱落)，";
                                break;
                            case "PL":
                                flawName += list[i] + "(破裂)，";
                                break;
                            case "CR":
                                flawName += list[i] + "(异物穿入)，";
                                break;
                            case "CJ":
                                flawName += list[i] + "(沉积)，";
                                break;
                            case "JG":
                                flawName += list[i] + "(结垢)，";
                                break;
                            case "ZW":
                                flawName += list[i] + "(障碍物)，";
                                break;
                            case "CQ":
                                flawName += list[i] + "(残墙、坝根)，";
                                break;
                            case "SG":
                                flawName += list[i] + "(树根)，";
                                break;
                            case "FZ":
                                flawName += list[i] + "(浮渣)，";
                                break;
                            default:
                                MessageBox.Show("检测类型出错");
                                break;
                        }
                        //flawName += list[i] + " ";
                    }
                    Image image_copy = image;
                    var graphics = Graphics.FromImage(image);
                    if (predictions.Count != 0)
                    {
                        image_copy.Save(image_handle_route + PythonCountHandle.ToString() + "_1.jpg");
                    }                                       
                    foreach (var prediction in predictions)      // iterate predictions to draw results
                    {
                        
                        double score = Math.Round(prediction.Score, 2);
                        graphics.DrawRectangles(new Pen(prediction.Label.Color, 5),
                        new[] { prediction.Rectangle });
                        var (x, y) = (prediction.Rectangle.X - 15, prediction.Rectangle.Y - 40);
                        graphics.DrawString($"{prediction.Label.Name} ({score})",
                        new Font("Consolas", 32, GraphicsUnit.Pixel), new SolidBrush(prediction.Label.Color),
                        new PointF(x, y));
                    }
                    if (list.Count != 0)
                    {
                        image.Save(image_handle_route + PythonCountHandle.ToString() + "_0.jpg");                        
                        string show_result = image_handle_route + PythonCountHandle.ToString() + "_0.jpg";
                        flaw_type.Text = flawName;
                        flaw_number.Text = flawNumber.ToString();
                        check_image.Text = PythonCountHandle.ToString() + "_0.jpg";
                        imageInteract = new imagePlayer(out_check.Size, show_result);
                        out_check.Controls.Clear();
                        out_check.Controls.Add(imageInteract);
                        imageInteract.newsize(out_check.Size);
                        PythonCountHandle++;
                    }
                    Python_Count++;
                }
            }
            //清理内存
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion
        #region   一键生成检测报告
        private void generate_report_Click(object sender, EventArgs e)
        {           
            if (Directory.GetDirectories(image_handle_route).Length > 0 || Directory.GetFiles(image_handle_route).Length > 0)
           {           
            word_app = new MSWord.ApplicationClass();
            wordDoc = word_app.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //wordDoc = word_app.Documents.Open(str_path, Missing.Value);
            word_app.Visible = true;
            wordPath = report_file + batchFileName +".doc";
            //下面是检测报告的封面设计
            
            //word_app.Selection.EndKey(ref unite, ref Nothing);    //将光标移动到文档末尾 
            //wordDoc.Content.InsertAfter("\n");
            //页面设置
            wordDoc.PageSetup.PaperSize = MSWord.WdPaperSize.wdPaperA4;              //设置纸张样式为A4纸
            wordDoc.PageSetup.Orientation = MSWord.WdOrientation.wdOrientPortrait;   //排列方式为垂直方向
            wordDoc.PageSetup.TopMargin = 57.0f;
            wordDoc.PageSetup.BottomMargin = 57.0f;
            wordDoc.PageSetup.LeftMargin = 57.0f;
            wordDoc.PageSetup.RightMargin = 57.0f;
            wordDoc.PageSetup.HeaderDistance = 30.0f;                                    //页眉位置

             //设置页眉
             word_app.ActiveWindow.View.Type = MSWord.WdViewType.wdNormalView;//普通视图（即页面视图）样式
             word_app.ActiveWindow.View.Type = MSWord.WdViewType.wdPrintView;
             word_app.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekPrimaryHeader;//进入页眉设置，其中页眉边距在页面设置中已完成
             word_app.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;//页眉中的文字居中对齐
             //word_app.ActiveWindow.ActivePane.Selection.InsertAfter("文档页眉");//在页眉的图片后面追加几个字

             word_app.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleNone;
             word_app.ActiveWindow.ActivePane.Selection.Borders[MSWord.WdBorderType.wdBorderBottom].Visible = false;
             word_app.ActiveWindow.View.Type = MSWord.WdViewType.wdPrintView;
             word_app.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;//退出页眉设置
             //插入页眉图片
             if(File.Exists((string)wordPath))                                        //测试阶段文件夹中只有一个文档
             { 
                File.Delete((string)wordPath);
             }
                //页码设置
             MSWord.PageNumbers pns = word_app.Selection.Sections[1].Headers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers;//获取当前页的号码
             pns.NumberStyle = MSWord.WdPageNumberStyle.wdPageNumberStyleNumberInDash;  //设置页码的风格，是Dash形还是圆形的
             pns.HeadingLevelForChapter = 0;
             pns.IncludeChapterNumber = false;
             pns.RestartNumberingAtSection = false;
             pns.StartingNumber = 0; //开始页页码？
             object pagenmbetal = MSWord.WdPageNumberAlignment.wdAlignPageNumberCenter; //将号码设置在中间
             object first = true;
             word_app.Selection.Sections[1].Footers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers.Add(ref pagenmbetal, ref first);

             //object WdLine2 = MSWord.WdUnits.wdLine;//换一行;  
             //wordApp.Selection.MoveDown(ref WdLine2, 6, ref Nothing);//向下跨15行输入表格，这样表格就在文字下方了，不过这是非主流的方法
             for (int i=0;i< PythonCountHandle; i++)     ///这里可能会有问题
            {
                AutoWord(i);
            }

            object format = MSWord.WdSaveFormat.wdFormatDocument;// office 2007就是wdFormatDocumentDefault
                                                                 //将wordDoc文档对象的内容保存为DOCX文档
            wordDoc.SaveAs(ref wordPath, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //看是不是要打印
            //wordDoc.PrintOut();
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象
            word_app.Quit(ref Nothing, ref Nothing, ref Nothing);
            //强制垃圾回收器回收资源
            GC.Collect();
            GC.WaitForPendingFinalizers();

            DeleteFolder(image_handle_route);
                //foreach (Process pro in wordProcess) //这里是找到那些没有界面的Word进程
                //    {
                //        IntPtr ip = pro.MainWindowHandle;

                //        string str = pro.MainWindowTitle; //发现程序中打开跟用户自己打开的区别就在这个属性
                //                                          //用户打开的str 是文件的名称，程序中打开的就是空字符串
                //        if (str == 文件名)
                //        {
                //            pro.Kill();
                //        }
                //    }
              foreach (System.Diagnostics.Process thisproc in System.Diagnostics.Process.GetProcessesByName("WINWORD"))  // 后台打开word文档后关闭
              {
                    thisproc.Kill();
              }
              foreach (System.Diagnostics.Process thisproc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))  // 后台打开excel文档后关闭
              {
                    thisproc.Kill();
              }
                //excel_app.Quit();
                MessageBox.Show("成功生成检测报告！！！");
                
           }
          else
          {
            MessageBox.Show("视频检测完成，不存在缺陷无报告生成！！！");
          }
        }         
        void AutoWord(int value)
        {        
            //word_app.Selection.GoTo(MSWord.WdGoToItem.wdGoToPage, MSWord.WdGoToDirection.wdGoToNext);  //将光标移动到下一行
            string wordFilepath;
            if(btn_batch.Checked)
            {
                wordFilepath = batchFileName;
            }
            else
            {
                wordFilepath = fileNameWithoutExtension;
            }
            MSExcel.Workbook workbook = excel_app.Workbooks.Open(report_file + wordFilepath + ".xlsx");
            MSExcel.Worksheet worksheet = workbook.Worksheets[1];
            wordDoc.Content.InsertAfter("\n");
            word_app.Selection.EndKey(ref unite, ref Nothing); //将光标移动到文档末尾               
            word_app.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;           
            
            //string cellValue = range.Value.ToString();

            int tableRow = 8;        //设置表格的行数和列数
            int tableColumn = 7;
            //定义一个word的表格对象
            MSWord.Table table = wordDoc.Tables.Add(word_app.Selection.Range,tableRow, tableColumn, ref Nothing, ref Nothing);

            //默认创建的表格没有边框，这里修改其属性，使得创建的表格带有边框 
            table.Borders.Enable = 1;//这个值可以设置得很大，例如5、13等等

            //表格的索引是从1开始的。  
            //wordDoc.Tables[1].Cell(1, 1).Range.Text = "录像文件";

            //下面为固定格式的设置
            //每一个空格填入的固定字符
            table.Cell(1, 1).Range.Text = "录像文件";
            table.Cell(2, 1).Range.Text = "管段类型";
            table.Cell(3, 1).Range.Text = "检测方向";
            table.Cell(4, 1).Range.Text = "检测地点 ";
            table.Cell(4, 4).Range.Text = "检测日期 ";
            table.Cell(5, 1).Range.Text = "距离(m) ";
            table.Cell(7, 1).Range.Text = "备注 ";
            table.Cell(5, 2).Range.Text = "缺陷名称代码";
            table.Cell(5, 3).Range.Text = "缺陷位置";
            table.Cell(5, 4).Range.Text = "分值";
            table.Cell(5, 5).Range.Text = "等级 ";
            table.Cell(5, 6).Range.Text = "管道内部状况描述 ";
            table.Cell(5, 7).Range.Text = "照片序号或说明";
            table.Cell(1, 4).Range.Text = "起始井号";
            table.Cell(2, 4).Range.Text = "管段材质";
            table.Cell(3, 4).Range.Text = "管段长度(m)";
            table.Cell(1, 6).Range.Text = "终止井号";
            table.Cell(2, 6).Range.Text = "管段直径(m)";
            table.Cell(3, 6).Range.Text = "检测长度(m)";

            //设置table样式
            table.Rows.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAtLeast;   //高度规则是：行高有最低值下限？
            table.Rows.Height = word_app.CentimetersToPoints(float.Parse("0.8"));          
            table.Range.Font.Size = 10.5F;
            table.Range.Font.Bold = 0;

            table.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;    //表格文本居中
            table.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;//文本垂直贴到居中
            table.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;                                 //表格位置居中
            //设置table边框样式
            table.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;  //表格外框是单线也可设置为别的双线
            table.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;   //表格内框是单线
            table.Rows[1].Range.Font.Bold = 0;//加粗
            table.Rows[1].Range.Font.Size = 10.5F;
            table.Cell(1, 1).Range.Font.Size = 10.5F;
            word_app.Selection.Cells.Height = 40;//所有单元格的高度        
            //除第一行外，其他行的行高都设置为20
            for (int i = 2; i <= tableRow; i++)
            {
                table.Rows[i].Height = 40;
            }
            table.Rows[6].Height = 60;

            //将表格左上角的单元格里的文字（“行” 和 “列”）居右
            //table.Cell(1, 1).Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
            //将表格左上角的单元格里面下面的“列”字移到左边，相比上一行就是将ParagraphFormat改成了Paragraphs[2].Format
            //table.Cell(1, 1).Range.Paragraphs[2].Format.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
            //table.Columns[1].Width = 65;//将第 1列宽度设置为50

            //将其他列的宽度都设置为75
            for (int i = 1; i <= tableColumn; i++)
            {
                table.Columns[i].Width = 65;
            }
            //添加表头斜线,并设置表头的样式
            //table.Cell(1, 1).Borders[MSWord.WdBorderType.wdBorderDiagonalDown].Visible = true;
            //table.Cell(1, 1).Borders[MSWord.WdBorderType.wdBorderDiagonalDown].Color = MSWord.WdColor.wdColorRed;
            //table.Cell(1, 1).Borders[MSWord.WdBorderType.wdBorderDiagonalDown].LineWidth = MSWord.WdLineWidth.wdLineWidth150pt;

            //最后一列添加图片
            table.Cell(8, 1).Merge(table.Cell(8, 7));
            //合并其他单元格
            table.Cell(1, 2).Merge(table.Cell(1, 3));   //横向合并
            table.Cell(2, 2).Merge(table.Cell(2, 3));
            table.Cell(3, 2).Merge(table.Cell(3, 3));
            table.Cell(4, 2).Merge(table.Cell(4, 3));
            table.Cell(4, 3).Merge(table.Cell(4, 5));
            table.Cell(7, 2).Merge(table.Cell(7, 7));
            table.Rows[8].Height = 370;              //设置新增加的这行表格的高度
            //向新添加的行的单元格中添加图片            
            string FileName = image_handle_route + value + "_0.jpg";      //检测后图片所在路径
            string fileNameCopy= image_handle_route + value + "_1.jpg";   //没有框的图片所在的路径
           
            int new_width = ((int)table.Cell(8, 1).Width);
            int new_height = ((int)table.Rows[8].Height);

            var newImage_0 = Image.FromFile(FileName);
            var newImage_1 = Image.FromFile(fileNameCopy);
            System.Drawing.Image img = new Bitmap(new_width, new_height);
            System.Drawing.Graphics newImageShow = System.Drawing.Graphics.FromImage(img);
            newImageShow.Clear(System.Drawing.Color.Transparent);
            newImageShow.DrawImage(newImage_1, 0, 0, newImage_1.Width, newImage_1.Height);
            Thread.Sleep(50);
            newImageShow.DrawImage(newImage_0, 0, 0, newImage_0.Width, newImage_0.Height);
            
            object LinkToFile = false;
            object SaveWithDocument = true;
            object Anchor = table.Cell(tableRow, 1).Range;      //选中要添加图片的单元格
          
            //先插入下方的未标注的图片
            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(fileNameCopy, ref LinkToFile, ref SaveWithDocument, ref Anchor);
            //由于是本文档的第1张图，所以这里是InlineShapes[1]
            MSWord.Shape s = wordDoc.Application.ActiveDocument.InlineShapes[1].ConvertToShape();
            s.WrapFormat.Type = MSWord.WdWrapType.wdWrapNone;
            table.Cell(8, 1).Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
            //wordDoc.Application.ActiveDocument.InlineShapes[1].Width =450;//图片宽度
            //wordDoc.Application.ActiveDocument.InlineShapes[1].Height = 280;//图片高度
            table.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;  //居中 

            //MSWord.Shape s = wordDoc.Application.ActiveDocument.InlineShapes[1].AlternativeText.
            //s.WrapFormat.Type = MSWord.WdWrapType.wdWrapSquare;
            //再插入已经标注的图片
            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Anchor);
            //由于是本文档的第2张图，所以这里是InlineShapes[2] 将图片设置为浮于文字上方
            MSWord.Shape s2 = wordDoc.Application.ActiveDocument.InlineShapes[1].ConvertToShape();
            s2.WrapFormat.Type = MSWord.WdWrapType.wdWrapNone;
            table.Cell(8, 1).Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
            table.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;  //居中
            //table.Cell(2, 3).Merge(table.Cell(4, 3));//纵向合并             
            try
            {
                //读取单元格数据
                MSExcel.Range cellvalue = worksheet.Cells[value + 2, 22];
                object range = cellvalue.Value2;
                string video_value = range.ToString();

                cellvalue= worksheet.Cells[value + 2, 2];
                range = cellvalue.Value2;
                string start_number = range.ToString();

                cellvalue = worksheet.Cells[value + 2, 3];
                range = cellvalue.Value2;
                string end_numder = range.ToString();

                cellvalue = worksheet.Cells[value + 2, 18];
                range = cellvalue.Value2;
                string pipe_mat2 = range.ToString();

                cellvalue = worksheet.Cells[value + 2, 6];
                range = cellvalue.Value2;
                string pipe_mat = range.ToString();
                cellvalue = worksheet.Cells[value + 2, 5];
                range = cellvalue.Value2;
                string pipe_diam = range.ToString();

                cellvalue = worksheet.Cells[value + 2, 17];
                range = cellvalue.Value2;
                string pipeLength = range.ToString();
                
                cellvalue = worksheet.Cells[value + 2, 13];
                range = cellvalue.Value2;
                string pipelocation = range.ToString();
               
                cellvalue = worksheet.Cells[value + 2, 16];
                range = cellvalue.Value2;
                string datetime1 = range.ToString();
               
                cellvalue = worksheet.Cells[value + 2, 17];
                range = cellvalue.Value2;
                string checkLength = range.ToString();
               
                cellvalue = worksheet.Cells[value + 2, 9];
                range = cellvalue.Value2;
                string flawName = range.ToString();
                
                cellvalue = worksheet.Cells[value + 2, 11];
                range = cellvalue.Value2;
                string flawLocation = range.ToString();
                
                cellvalue = worksheet.Cells[value + 2, 12];
                range = cellvalue.Value2;
                string flawLevel = range.ToString();
                
                cellvalue = worksheet.Cells[value + 2, 7];
                range = cellvalue.Value2;
                string photoNumber = range.ToString();
                //固定位置填上相应的参数
                table.Cell(1, 2).Range.Text = video_value;
                table.Cell(1, 4).Range.Text = start_number;
                table.Cell(1, 6).Range.Text = end_numder;
                table.Cell(2, 2).Range.Text = pipe_mat2;
                table.Cell(2, 4).Range.Text = pipe_mat;
                //table.Cell(2, 5).Range.Text = pipeMat.Text;
                table.Cell(2, 6).Range.Text = pipe_diam;
                table.Cell(3, 4).Range.Text = pipeLength;
                table.Cell(3, 6).Range.Text = checkLength;
                table.Cell(4, 2).Range.Text = pipelocation;
                string nowTime = DateTime.Now.ToString();
                table.Cell(4, 4).Range.Text = nowTime;
                table.Cell(6, 1).Range.Text = checkLength;
                table.Cell(6, 2).Range.Text = flawName;
                table.Cell(6, 3).Range.Text = flawLocation;
                table.Cell(6, 7).Range.Text = photoNumber;
                table.Cell(7, 2).Range.Text = "/";

                //出问题了就要将excel关闭
                //wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
                //关闭wordApp组件对象
                //excel_app.Quit(ref Nothing, ref Nothing, ref Nothing);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch( Exception e)
            {
                //释放com对象
                //Marshal.ReleaseComObject(range);
                //Marshal.ReleaseComObject(worksheet);
                //Marshal.ReleaseComObject(workbook);
                //Marshal.ReleaseComObject(excelApp);
                //GC.Collect();
                workbook.Close();
                GC.WaitForPendingFinalizers();
                MessageBox.Show(e.ToString());
            }
            //获取excel表格中的数据
            //string video_value = worksheet.Cells[value + 2, 8];          
            //string start_number = worksheet.Cells[value + 2, 2];
            //string end_numder = worksheet.Cells[value + 2, 3];
            //string pipe_mat2 = worksheet.Cells[value + 2, 18];
            //string pipe_mat = worksheet.Cells[value + 2, 6];
            //string pipe_diam = worksheet.Cells[value + 2, 5];
            //string pipeLength = worksheet.Cells[value + 2, 17];
            //string pipelocation = worksheet.Cells[value + 2, 17];
            //string datetime1 = worksheet.Cells[value + 2, 16];
            //string checkLength = worksheet.Cells[value + 2, 17];
            //string flawName = worksheet.Cells[value + 2, 9];
            //string flawLocation = worksheet.Cells[value + 2, 11];
            //string flawLevel = worksheet.Cells[value + 2, 12];
            //string photoNumber = worksheet.Cells[value + 2, 7];         
            word_app.Selection.EndKey(ref unite, ref Nothing); //将光标移动到文档末尾         
            wordDoc.Content.InsertAfter("\n");
            //wordDoc.Content.InsertAfter(DateTime.Now.ToShortTimeString());   //文档末尾插入时间
            workbook.Close();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //WdSaveFormat为Word 2003文档的保存格式
        }
        #endregion
    }
}
