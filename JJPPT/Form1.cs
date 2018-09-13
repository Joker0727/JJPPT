using HtmlAgilityPack;
using MyHttp;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using SqlLiteHelperDemo;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JJPPT
{
    public partial class Form1 : Form
    {
        private string url = "https://jjppt.pc1000000.cn/";
        private string pageUrlPrefix = "https://jjppt.pc1000000.cn/ppt?page=";
        private int totalPages = 0;
        public string basePath = AppDomain.CurrentDomain.BaseDirectory;
        public string sqlitePath = AppDomain.CurrentDomain.BaseDirectory + @"sqlite3.db";
        public SQLiteHelper sqlLiteHelper = null;
        public SeleniumHelper sel = null;
        public int IsDownLoadedUrlCount = 0;
        public bool IsOk = true;
        public string downLoadPath = string.Empty;
        public string movePath = string.Empty;
        public string defaultPath = AppDomain.CurrentDomain.BaseDirectory + "DefaultPath.ini";
        public string workId = "ww-0005";
        Thread th = null;
        Thread downLoadPPTThread = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.MaximizeBox = false;
            if (!File.Exists(sqlitePath))
            {
                IsOk = false;
                MessageBox.Show("数据故障！", "JJPPT");
                return;
            }
            try
            {
                sqlLiteHelper = new SQLiteHelper(sqlitePath);
            }
            catch (Exception ex)
            {
                IsOk = false;
                MessageBox.Show(ex.ToString());
                WriteLog(ex.ToString());
                return;
            }
            InitPath();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string btnStr = this.button1.Text.Trim();
            if (btnStr == "下 载 链 接")
            {
                if (!IsOk)
                {
                    MessageBox.Show("数据故障！", "JJPPT");
                    return;
                }
                //if (!IsAuthorised(workId))
                //{
                //    MessageBox.Show("请检查网络！", "JJPPT");
                //    return;
                //}
                ToManageThread();
                this.button1.Text = "暂 停";
            }
            else
            {
                this.button1.Text = "下 载 链 接";
                if (th != null)
                {
                    th.Abort();
                    th = null;
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string btnStr = this.button2.Text.Trim();
            if (btnStr == "下 载 PPT")
            {
                CloseFireFoxAndGeckodriver();
                if (!IsOk)
                {
                    MessageBox.Show("数据故障！", "JJPPT");
                    return;
                }

                downLoadPath = this.textBox1.Text;
                movePath = this.textBox2.Text;

                if (string.IsNullOrEmpty(downLoadPath) || string.IsNullOrEmpty(movePath))
                {
                    MessageBox.Show("路径不能为空！", "JJPPT");
                    return;
                }
                //if (!IsAuthorised(workId))
                //{
                //    MessageBox.Show("请检查网络！", "JJPPT");
                //    return;
                //}
                File.WriteAllText(defaultPath, downLoadPath + "\r\n" + movePath);

                sel = new SeleniumHelper(1);

                sel.driver.Navigate().GoToUrl(url);
                MessageBoxButtons message = MessageBoxButtons.OKCancel;
                DialogResult dr = MessageBox.Show("请先登录成功后，再点击确定！", "JJPPT", message);
                if (dr == DialogResult.OK)
                {
                    downLoadPPTThread = new Thread(DownLoadPPT);
                    downLoadPPTThread.IsBackground = true;
                    downLoadPPTThread.Start();
                    this.button2.Text = "暂 停";
                }
            }
            else
            {
                this.button2.Text = "下 载 PPT";
                if (downLoadPPTThread != null)
                {
                    downLoadPPTThread.Abort();
                    downLoadPPTThread = null;
                }
                CloseFireFoxAndGeckodriver();
            }
        }
        private void textBox1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = dialog.SelectedPath;
            }
        }
        private void textBox2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.textBox2.Text = dialog.SelectedPath;
            }
        }
        /// <summary>
        /// 多线程管理/
        /// </summary>
        public void ToManageThread()
        {
            try
            {
                totalPages = GetTotalPages() * 32;
                if (totalPages > 0)
                {
                    th = new Thread(CyclicDownload);
                    th.IsBackground = true;
                    th.Start();
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }
        }
        /// <summary>
        /// 循环下载操作
        /// </summary>
        public void CyclicDownload()
        {
            string pageUrl = string.Empty;
            for (int i = 1; i < totalPages; i++)
            {
                try
                {
                    pageUrl = pageUrlPrefix + i;
                    DownLoadUrl(pageUrl);
                }
                catch (Exception ex)
                {
                    WriteLog(ex.ToString());
                }
            }
        }
        /// <summary>
        /// 下载url
        /// </summary>
        public void DownLoadUrl(string pageUrl)
        {
            string pptId, pptName, tempUrl, sqlStr, pptUrl = string.Empty;

            try
            {
                string pageHtmlStr = GetHtml(pageUrl);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(pageHtmlStr);
                HtmlNodeCollection h4aList = doc.DocumentNode.SelectNodes("//ul[@class='posts clearfix']/li[@class='postsItem']/div[@class='p_shadow']/h4[@class='imgTitle']/a");

                foreach (var h4a in h4aList)
                {
                    try
                    {
                        pptName = h4a.InnerText;
                        tempUrl = h4a.GetAttributeValue("href", "");
                        if (tempUrl.Contains(url))
                            pptId = tempUrl.Replace("https://jjppt.pc1000000.cn/d/", "");
                        else
                            pptId = tempUrl.Replace("/d/", "");

                        if (IsNumeric(pptId))
                        {
                            pptUrl = "https://jjppt.pc1000000.cn/vip/download?id=" + pptId + "&type=1";
                            sqlStr = @"INSERT INTO JJPPTtable (PPTId,PPTName,PPTUrl,IsDownLoad) 
                                VALUES (" + pptId + ",'" + pptName + "','" + pptUrl + "', 0); ";
                            sqlLiteHelper.RunSql(sqlStr);
                            IsDownLoadedUrlCount++;
                            this.label2.Invoke(new Action(() =>
                                {
                                    this.label2.Text = IsDownLoadedUrlCount.ToString() + "/" + totalPages;
                                }));

                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLog(ex.ToString());
                    }
                }
            }
            catch (Exception er)
            {
                WriteLog(er.ToString());
            }
        }
        /// <summary>
        /// 获取总页数
        /// </summary>
        public int GetTotalPages()
        {
            int count = 0;
            string mainStr = string.Empty;

            try
            {
                mainStr = GetHtml(url);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(mainStr);
                HtmlNode totalPageNodeA = doc.DocumentNode.SelectNodes("//div[@class='trunPage']/a")[3];
                count = IsNumeric(totalPageNodeA.InnerText) ? int.Parse(totalPageNodeA.InnerText) : 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                WriteLog(ex.ToString());
            }

            return count;
        }
        /// <summary>
        /// 下载ppt
        /// </summary>
        public void DownLoadPPT()
        {
            List<PPTDto> pptDtoList = GetPPTDetails();
            int totalPPT = pptDtoList.Count();
            int downLoadPPTCount = 0;

            foreach (var pptDto in pptDtoList)
            {
                try
                {
                    sel.driver.Navigate().GoToUrl(pptDto.PPTUrl);
                    ((IJavaScriptExecutor)sel.driver).ExecuteScript("location.reload()");
                    Task task = Task.Run(() => ChangeNameAndMoveFile(pptDto.PPTName));
                    task.Wait();

                    UpDataState(pptDto.PPTId);
                    downLoadPPTCount++;
                    this.label4.Invoke(new Action(() =>
                    {
                        this.label4.Text = downLoadPPTCount + "/" + totalPPT;
                    }));
                }
                catch (Exception ex)
                {
                    WriteLog(ex.ToString());
                }
            }
            CloseFireFoxAndGeckodriver();
        }
        /// <summary>
        /// 修改文件路径并改名
        /// </summary>
        /// <param name="pptName"></param>
        public void ChangeNameAndMoveFile(string pptName)
        {
            string newPath = string.Empty;
            bool flag = true;
            while (flag)
            {
                Thread.Sleep(1000);
                if (Directory.Exists(downLoadPath))
                {
                    DirectoryInfo theFolder = new DirectoryInfo(downLoadPath);
                    foreach (FileInfo file in theFolder.GetFiles())
                    {
                        try
                        {
                            if (file.Extension.ToLower() == ".ppt" ||
                                file.Extension.ToLower() == ".pptx" ||
                                file.Extension.ToLower() == ".rar")
                            {
                                newPath = movePath + @"\" + pptName + file.Extension;
                                Thread.Sleep(1000);
                                if (File.Exists(newPath))
                                    File.Delete(newPath);
                                File.Move(file.FullName, newPath);
                                flag = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteLog(ex.ToString());
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 获取ppt信息集合
        /// </summary>
        /// <returns></returns>
        public List<PPTDto> GetPPTDetails()
        {
            string sqlStr = "SELECT * FROM JJPPTtable WHERE IsDownLoad = 0";
            List<PPTDto> pptDtoList = new List<PPTDto>();

            try
            {
                DataTable pptDataTable = sqlLiteHelper.SearchSql(sqlStr);

                pptDtoList = ModelHelper<PPTDto>.DataTableToModel(pptDataTable);
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }

            return pptDtoList;
        }
        /// <summary>
        /// 更新PPT下载状态
        /// </summary>
        /// <param name="pptId"></param>
        /// <returns></returns>
        public bool UpDataState(long pptId)
        {
            bool result = false;
            string sqlStr = "UPDATE JJPPTtable SET IsDownLoad = 1 WHERE PPTId = " + pptId + "";

            try
            {
                result = sqlLiteHelper.RunSql(sqlStr);
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }

            return result;
        }
        /// <summary>
        /// 关闭火狐浏览器和geckodriver驱动
        /// </summary>
        public void CloseFireFoxAndGeckodriver()
        {
            KillProcess("firefox");
            KillProcess("chrome");
            KillProcess("geckodriver");
            KillProcess("chromedriver");
        }
        /// <summary>
        /// 杀死进程
        /// </summary>
        /// <param name="pName">进程名</param>
        public void KillProcess(string pName)
        {
            Process[] process;//创建一个PROCESS类数组
            process = Process.GetProcesses();//获取当前任务管理器所有运行中程序
            foreach (Process proces in process)//遍历
            {
                try
                {
                    if (proces.ProcessName == pName)
                    {
                        proces.Kill();
                    }
                }
                catch (Exception ex) { }
            }
        }
        /// <summary>
        /// 判断字符串是不是数字类型的 true是数字
        /// </summary>
        /// <param name="value">需要检测的字符串</param>
        /// <returns>true是数字</returns>
        public bool IsNumeric(string value)
        {
            return Regex.IsMatch(value, @"^\d(\.\d+)?|[1-9]\d+(\.\d+)?$");
        }
        /// <summary>
        /// 日志打印
        /// </summary>
        /// <param name="log"></param>
        public void WriteLog(string log)
        {
            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + "log\\";//日志文件夹
                DirectoryInfo dir = new DirectoryInfo(path);
                if (!dir.Exists)//判断文件夹是否存在
                    dir.Create();//不存在则创建

                FileInfo[] subFiles = dir.GetFiles();//获取该文件夹下的所有文件
                foreach (FileInfo f in subFiles)
                {
                    string fname = Path.GetFileNameWithoutExtension(f.FullName); //获取文件名，没有后缀
                    DateTime start = Convert.ToDateTime(fname);//文件名转换成时间
                    DateTime end = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd"));//获取当前日期
                    TimeSpan sp = end.Subtract(start);//计算时间差
                    if (sp.Days > 30)//大于30天删除
                        f.Delete();
                }

                string logName = DateTime.Now.ToString("yyyy-MM-dd") + ".log";//日志文件名称，按照当天的日期命名
                string fullPath = path + logName;//日志文件的完整路径
                string contents = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " -> " + log + "\r\n";//日志内容

                File.AppendAllText(fullPath, contents, Encoding.UTF8);//追加日志

            }
            catch (Exception ex)
            {

            }
        }
        /// <summary>
        /// 获取html源码
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public string GetHtml(string url)
        {
            string strHTML = string.Empty;

            try
            {
                WebClient myWebClient = new WebClient();
                Stream myStream = myWebClient.OpenRead(url);
                StreamReader sr = new StreamReader(myStream, Encoding.GetEncoding("utf-8"));
                strHTML = sr.ReadToEnd();
                myStream.Close();
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }

            return strHTML;
        }
        /// <summary>
        /// 初始化路径
        /// </summary>
        public void InitPath()
        {
            string defaultPathStr = File.ReadAllText(defaultPath);
            string[] pathArr = Regex.Split(defaultPathStr, "\r\n", RegexOptions.IgnoreCase);
            if (pathArr.Count() == 2)
            {
                this.textBox1.Text = pathArr[0];
                this.textBox2.Text = pathArr[1];
            }
        }
        /// <summary>
        /// 授权
        /// </summary>
        /// <param name="workId"></param>
        /// <returns></returns>
        public bool IsAuthorised(string workId)
        {
            string conStr = "Server=111.230.149.80;DataBase=MyDB;uid=sa;pwd=1add1&one";
            using (SqlConnection con = new SqlConnection(conStr))
            {
                string sql = string.Format("select count(*) from MyWork Where PassState = 1 and WorkId ='{0}'", workId);
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    con.Open();
                    int count = int.Parse(cmd.ExecuteScalar().ToString());
                    if (count > 0)
                        return true;
                }
            }
            return false;
        }
        /// <summary>
        /// 判断是否超过100
        /// </summary>
        /// <returns></returns>
        public bool DownLoadCount()
        {
            bool result = false;
            string sqlStr = "select count(*) from JJPPTtable where IsDownLoad = 1";
            try
            {
                string objStr = sqlLiteHelper.GetScalar(sqlStr).ToString();
                int currentCount = IsNumeric(objStr) ? int.Parse(objStr) : 100;
                if (currentCount < 100)
                {
                    result = true;
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }
            return result;
        }
    }
}
