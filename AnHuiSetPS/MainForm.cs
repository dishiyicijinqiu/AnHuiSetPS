using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Windows.Forms;
using TrainTicketLogin.Helper;

namespace AnHuiSetPS
{
    public partial class MainForm : Form
    {
        CookieCollection cookies = new CookieCollection();
        BindingSource bs = new BindingSource();
        string UnAddSql = "SELECT * FROM [Sheet1$] WHERE 是否添加=@是否添加 and 产品流水号 is not null";
        string CompleteSql = "UPDATE [Sheet1$] SET 是否添加=@是否添加 WHERE 产品流水号=@产品流水号";
        /// <summary>
        /// 所在地区列表
        /// </summary>
        Dictionary<string, string> CityList = new Dictionary<string, string>();
        /// <summary>
        /// 医疗名称列表
        /// </summary>
        Dictionary<string, string> HospitalList = new Dictionary<string, string>();
        /// <summary>
        /// 配送企业列表
        /// </summary>
        Dictionary<string, string> CompanyList = new Dictionary<string, string>();
        private bool IsStop = true;
        private bool IsLogin = false;
        object lockobj = new object();
        private string loginName = string.Empty;
        MyWorkThread[] threads = new MyWorkThread[Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["ThreadNum"])];
        Stack<DataRow> stack = new Stack<DataRow>();
        #region 构造函数及载入函数
        public MainForm()
        {
            InitializeComponent();
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            this.IsLogin = HasLogin();
            if (!this.IsLogin)
                GetCode();
            else
            {
                this.Text = string.Format("安徽设置配送，当前登录账号为：{0}", this.loginName);
            }
            this.cbbUsers.SelectedIndex = 1;
#if DEBUG
            this.btnQuery.Visible = true;
#else 
            this.btnQuery.Visible = false;
#endif
        }

        private bool HasLogin()
        {
            var LastCookie = GetCookie();
            if (LastCookie == null)
                return false;
            if (LastCookie != null)
            {
                cookies.Add(LastCookie);
            }
            string url = "http://jy.ahyycg.cn:8080/Default.aspx";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Accept = "text/html, application/xhtml+xml, */*";
            request.AllowAutoRedirect = false;
            request.AutomaticDecompression = DecompressionMethods.GZip;
            request.Headers["Accept-Language"] = "zh-CH;en-US";
            request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
            request.KeepAlive = true;
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
            request.Timeout = 50000;
            request.Method = "GET";
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(cookies);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            UpdateLocalCookies(response.Cookies);
            var myResponseStream = response.GetResponseStream();
            var myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
            string outdata = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            response.Close();
            string strPattner = "<img alt=\"\" src=\"./lib/images/smile.gif\"/> .*好！(?<value>.*?)</a> ";
            var regex = new Regex(strPattner);
            var match = regex.Match(outdata);
            if (match == null)
                return false;
            if (match.Groups.Count < 2)
                return false;
            loginName = match.Groups["value"].Value;
            return true;
        }
        #endregion

        private void btnReGetCode_Click(object sender, EventArgs e)
        {
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            try
            {
                HttpWebRequest request = null;
                string url = "http://jy.ahyycg.cn:8080/UserLogin.aspx";   //登录页面
                request = (HttpWebRequest)WebRequest.Create(url);
                request.Accept = "text/html, application/xhtml+xml, */*";
                request.AllowAutoRedirect = false;
                request.AutomaticDecompression = DecompressionMethods.GZip;
                request.Headers["Accept-Language"] = "zh-CH;en-US";
                request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
                request.KeepAlive = true;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                request.Timeout = 50000;
                request.Referer = "http://jy.ahyycg.cn:8080/UserLogin.aspx";
                request.Method = "POST";
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
                StringBuilder postStrSbuilder = new StringBuilder();
                postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTTARGET"), HttpUtility.UrlEncode(string.Empty));
                postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTARGUMENT"), HttpUtility.UrlEncode(string.Empty));
                postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATE"), HttpUtility.UrlEncode("/wEPDwUJNjMyODYwNzYzZGQtMroG+Ni2RimNxdoNmNILsFesPw=="));
                postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATEGENERATOR"), HttpUtility.UrlEncode("7A1355CA"));
                postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("txtUserId"), HttpUtility.UrlEncode(txtUserName.Text));
                postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("txtUserPwd"), HttpUtility.UrlEncode(txtPassword.Text));
                postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("txtCode"), HttpUtility.UrlEncode(txtCode.Text));
                postStrSbuilder.AppendFormat("{0}={1}", HttpUtility.UrlEncode("btnLogin"), HttpUtility.UrlEncode("登 录"));
                Stream myRequestStream = request.GetRequestStream();
                request.ContentType = "application/x-www-form-urlencoded";
                StreamWriter myStreamWriter = new StreamWriter(myRequestStream, Encoding.UTF8);
                myStreamWriter.Write(postStrSbuilder.ToString());
                myStreamWriter.Close();
                myRequestStream.Close();
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                UpdateLocalCookies(response.Cookies);
                Stream myResponseStream = response.GetResponseStream();
                StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
                string outdata = myStreamReader.ReadToEnd();
                myStreamReader.Close();
                myResponseStream.Close();
                if (outdata.Contains("Default.aspx"))
                {
                    foreach (Cookie cookie in cookies)
                    {
                        API.InternetSetCookie("https://" + cookie.Domain.ToString() + ":8080", cookie.Name.ToString(), cookie.Value.ToString() + ";expires=Sun,22-Feb-2099 00:00:00 GMT");
                    }
                    this.IsLogin = this.HasLogin();
                    this.Text = string.Format("安徽设置配送，当前登录账号为：{0}", this.loginName);
                }
                else
                {
                    if (outdata.Contains("验证码不正确！"))
                    {
                        MessageBoxEx.Show(this, "验证码不正确！");
                    }
                    else if (outdata.Contains("用户名密码不匹配！"))
                    {
                        MessageBoxEx.Show(this, "用户名密码不匹配！");
                    }
                    else
                    {
                        MessageBoxEx.Show(this, "未知错误，登陆失败");
                    }
                }
            }
            catch (Exception ex)
            {
                OnError("登录", ex.Message);
            }
        }

        private void btnSetFile_Click(object sender, EventArgs e)
        {
            try
            {

                OpenFileDialog dia = new OpenFileDialog();
                dia.Multiselect = false;
                dia.Filter = "(*.xls,*.xlsx)|*.xls;*.xlsx";
                if (dia.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                {
                    this.tbPath.Text = dia.FileName;
                    ExcelHelper.fileName = dia.FileName;
                    var ds = ExcelHelper.GetReader(UnAddSql, new OleDbParameter[] { new OleDbParameter("@是否添加", "否") });
                    bs.DataSource = ds.Tables[0];
                    this.dgv.DataSource = bs;
                }
            }
            catch (Exception ex)
            {
                OnError("打开数据文件", ex.Message);
            }
        }

        private void btnOpenWebSite_Click(object sender, EventArgs e)
        {
            this.IsLogin = this.HasLogin();
            if (!this.IsLogin)
            {
                this.Text = "安徽设置配送";
                MessageBoxEx.Show(this, "失去连接请重新登陆");
                return;
            }
            foreach (Cookie cookie in cookies)
            {
                if (!API.InternetSetCookie("https://" + cookie.Domain.ToString() + ":8080", cookie.Name.ToString(), cookie.Value.ToString() + ";expires=Sun,22-Feb-2099 00:00:00 GMT"))
                {
                    MessageBoxEx.Show(this, "失去连接请重新登陆");
                }
            }
            Process.Start("IExplore.exe", "http://jy.ahyycg.cn:8080/Default.aspx");
        }

        private void btnSetPS_Click(object sender, EventArgs e)
        {
            if (!IsLogin)
            {
                MessageBoxEx.Show(this, "请先登录");
                return;
            }
            this.IsLogin = this.HasLogin();
            if (!this.IsLogin)
            {
                this.Text = "安徽设置配送";
                MessageBoxEx.Show(this, "失去连接请重新登陆");
                return;
            }
            DataTable dt = this.bs.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBoxEx.Show(this, "当前未选择任何数据，或者选择数据为0");
                return;
            }
            this.IsStop = false;
            this.SetEnable(false);
            stack.Clear();
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                stack.Push(dt.Rows[i]);
            }
            for (int i = 0; i < threads.Length; i++)
            {
                threads[i] = new MyWorkThread(new Thread(new ThreadStart(StartPSEntity)));
                threads[i].IsBackground = true;
                threads[i].Start();
            }
        }

        private void StartPSEntity()
        {
            while (true)
            {
                if (this.IsStop)
                {
                    this.IsStop = true;
                    foreach (MyWorkThread thread in threads)
                    {
                        if (thread.WorkThreadId == Thread.CurrentThread.ManagedThreadId)
                            thread.WorkComplete = true;
                    }
                    this.OnEnd();
                    break;
                }
                DataRow dr = null;
                lock (lockobj)
                {
                    if (stack.Count > 0)
                        dr = stack.Pop();
                }
                if (dr == null)
                {
                    this.IsStop = true;
                    foreach (MyWorkThread thread in threads)
                    {
                        if (thread.WorkThreadId == Thread.CurrentThread.ManagedThreadId)
                            thread.WorkComplete = true;
                    }
                    this.OnEnd();
                    break;
                }
                Entity entity = new Entity();
                entity.产品流水号 = dr["产品流水号"].ToString();
                entity.医疗名称 = dr["医疗名称"].ToString();
                entity.所在地区 = dr["所在地区"].ToString();
                entity.是否添加 = dr["是否添加"].ToString();
                entity.配送企业 = dr["配送企业"].ToString();
                try
                {
                    var result = SetPS(entity);
                    Complete(entity.产品流水号);
                    lock (lockobj)
                    {
                        this.Invoke(new MethodInvoker(() =>
                        {
                            int entityIndex = this.bs.Find("产品流水号", entity.产品流水号);
                            this.bs.RemoveAt(entityIndex);
                            this.listBox1.Items.Add(string.Format("{0}:{1}", entity.产品流水号, result));
                        }));
                    }
                }
                catch (Exception ex)
                {
                    OnError("设置配送", entity.产品流水号 + ":" + ex.Message);
                }
            }
        }

        private void OnEnd()
        {
            lock (lockobj)
            {
                if (!this.IsStop)
                    return;
                foreach (MyWorkThread thread in threads)
                {
                    if (thread.WorkThreadId == Thread.CurrentThread.ManagedThreadId)
                    {
                        continue;
                    }
                    if (!thread.WorkComplete)
                        return;
                }
                this.SetEnable(true);
                this.Invoke(new MethodInvoker(() =>
                {
                    MessageBoxEx.Show(this, "全部处理完成");
                }));
            }
        }

        private string SetPS(Entity entity)
        {
            string ViewStatePattner = "id=\"__VIEWSTATE\" value=\"(?<value>.*?)\"";
            string SelectPattner = "<option value=\"(?<key>.*?)\">(?<value>.*?)</option>";
            string url = string.Empty;
            string str__EVENTARGUMENT = string.Empty;
            string str__LASTFOCUS = string.Empty;
            string strhfdCompanyName = string.Empty;
            string str__VIEWSTATEGENERATOR = string.Empty;
            string strAspNetPager1_input = string.Empty;//当前第几页
            string strAspNetPager1_pagesize = string.Empty;//每页大小
            string str__EVENTTARGET = string.Empty;
            string str__VIEWSTATE = string.Empty;//ViewState
            string strddlCity = string.Empty;//所在地区
            string strddlhosname = string.Empty;//医疗名称
            string strcompanyId = string.Empty;//配送企业
            Stream myResponseStream;
            StreamReader myStreamReader;
            HttpWebResponse response;
            HttpWebRequest request;
            string outdata4 = string.Empty;
            #region Get请求
            url = string.Format("http://jy.ahyycg.cn:8080/Enterprise/RelationQuery/ProductCompanySubArea.aspx?PID={0}&returnUrl=/Enterprise/RelationQuery/RelationQueryUnpack.aspx", entity.产品流水号);
            request = (HttpWebRequest)WebRequest.Create(url);
            request.Accept = "text/html, application/xhtml+xml, */*";
            request.AllowAutoRedirect = false;
            request.AutomaticDecompression = DecompressionMethods.GZip;
            request.Headers["Accept-Language"] = "zh-CH;en-US";
            request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
            request.KeepAlive = true;
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
            request.Timeout = 50000;
            request.Method = "GET";
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(cookies);
            response = (HttpWebResponse)request.GetResponse();
            UpdateLocalCookies(response.Cookies);
            myResponseStream = response.GetResponseStream();
            myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
            string outdata = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            response.Close();
            #region 获取所在地区列表
            if (CityList.Count <= 0)
            {
                lock (lockobj)
                {
                    if (CityList.Count <= 0)
                    {
                        string startString = "<select name=\"ctl00$ContentPlaceHolder1$ddlCity\"";
                        string endString = "</select>";
                        int StartIndex = outdata.IndexOf(startString);
                        var Tempoutdata = outdata.Substring(StartIndex);
                        int EndIndex = Tempoutdata.IndexOf(endString);
                        Tempoutdata = Tempoutdata.Substring(0, EndIndex);
                        var regex = new Regex(SelectPattner);
                        var matchs = regex.Matches(Tempoutdata);
                        CityList.Clear();
                        foreach (Match match in matchs)
                        {
                            CityList.Add(match.Groups["value"].Value.Trim(), match.Groups["key"].Value.Trim());
                        }
                    }
                }
            }
            #endregion
            #region 医疗名称列表
            if (HospitalList.Count <= 0)
            {
                lock (lockobj)
                {
                    if (HospitalList.Count <= 0)
                    {
                        string startString = "<select name=\"ctl00$ContentPlaceHolder1$ddlhosname\"";
                        string endString = "</select>";
                        int StartIndex = outdata.IndexOf(startString);
                        string Tempoutdata = outdata.Substring(StartIndex);
                        int EndIndex = Tempoutdata.IndexOf(endString);
                        Tempoutdata = Tempoutdata.Substring(0, EndIndex);
                        var regex = new Regex(SelectPattner);
                        var matchs = regex.Matches(Tempoutdata);
                        HospitalList.Clear();
                        foreach (Match matchItem in matchs)
                        {
                            HospitalList.Add(matchItem.Groups["value"].Value.Trim(), matchItem.Groups["key"].Value.Trim());
                        }
                    }
                }
            }
            #endregion
            #endregion
            if (CompanyList.Count <= 0)
            {
                #region 获取配送企业列表
                lock (lockobj)
                {
                    if (CompanyList.Count <= 0)
                    {
                        string comIdPattner = "var comID = '(?<value>.*?)';";
                        string ComNamePattner = "var comName = '(?<value>.*?)';";
                        var regex = new Regex(comIdPattner);
                        var match = regex.Match(outdata);
                        string comId = match.Groups["value"].Value;
                        regex = new Regex(ComNamePattner);
                        match = regex.Match(outdata);
                        string ComName = match.Groups["value"].Value;
                        url = string.Format("http://jy.ahyycg.cn:8080/Enterprise/RelationQuery/CompanySelect.aspx?id={0}&name={1}", comId, ComName);
                        #region 请求Get请求
                        request = (HttpWebRequest)WebRequest.Create(url);
                        request.Accept = "text/html, application/xhtml+xml, */*";
                        request.AllowAutoRedirect = false;
                        request.AutomaticDecompression = DecompressionMethods.GZip;
                        request.Headers["Accept-Language"] = "zh-CH;en-US";
                        request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
                        request.KeepAlive = true;
                        request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                        request.Timeout = 50000;
                        request.Method = "GET";
                        request.CookieContainer = new CookieContainer();
                        request.CookieContainer.Add(cookies);
                        response = (HttpWebResponse)request.GetResponse();
                        UpdateLocalCookies(response.Cookies);
                        myResponseStream = response.GetResponseStream();
                        myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
                        string outdata1 = myStreamReader.ReadToEnd();
                        myStreamReader.Close();
                        myResponseStream.Close();
                        response.Close();
                        #endregion
                        #region Post请求，获取配送企业列表
                        request = (HttpWebRequest)WebRequest.Create(url);
                        request.Headers["Pragma"] = "no-cache";
                        request.Accept = "text/html, application/xhtml+xml, */*";
                        request.AllowAutoRedirect = false;
                        request.AutomaticDecompression = DecompressionMethods.GZip;
                        request.Headers["Accept-Language"] = "zh-CH;en-US";
                        request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
                        request.KeepAlive = true;
                        request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                        request.Timeout = 50000;
                        request.Method = "POST";
                        request.CookieContainer = new CookieContainer();
                        request.CookieContainer.Add(cookies);
                        strAspNetPager1_pagesize = int.MaxValue.ToString();
                        request.CookieContainer.Add(new Cookie("5", strAspNetPager1_pagesize, "/", "jy.ahyycg.cn"));
                        request.Referer = url;
                        StringBuilder postStrSbuilder = new StringBuilder();
                        regex = new Regex(ViewStatePattner);
                        match = regex.Match(outdata1);
                        if (match != null)
                            str__VIEWSTATE = match.Groups["value"].Value;
                        str__VIEWSTATEGENERATOR = "AD8CA982";
                        str__EVENTTARGET = "AspNetPager1";
                        str__EVENTARGUMENT = string.Empty;
                        string strtxtProduceCompany_PS = string.Empty;
                        strAspNetPager1_input = "1";
                        postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATE"), HttpUtility.UrlEncode(str__VIEWSTATE));
                        postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATEGENERATOR"), HttpUtility.UrlEncode(str__VIEWSTATEGENERATOR));
                        postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTTARGET"), HttpUtility.UrlEncode(str__EVENTTARGET));
                        postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTARGUMENT"), HttpUtility.UrlEncode(str__EVENTARGUMENT));
                        postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("txtProduceCompany_PS"), HttpUtility.UrlEncode(strtxtProduceCompany_PS));
                        postStrSbuilder.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$AspNetPager1_input"), HttpUtility.UrlEncode(strAspNetPager1_input));
                        postStrSbuilder.AppendFormat("{0}={1}", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$AspNetPager1_pagesize"), HttpUtility.UrlEncode(strAspNetPager1_pagesize));
                        byte[] postBytes = Encoding.UTF8.GetBytes(postStrSbuilder.ToString());
                        request.ContentType = "application/x-www-form-urlencoded";
                        request.ContentLength = postBytes.Length;
                        Stream postDataStream = request.GetRequestStream();
                        postDataStream.Write(postBytes, 0, postBytes.Length);
                        postDataStream.Close();
                        postDataStream.Dispose();
                        response = (HttpWebResponse)request.GetResponse();
                        UpdateLocalCookies(response.Cookies);
                        myResponseStream = response.GetResponseStream();
                        myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
                        string outdata2 = myStreamReader.ReadToEnd();
                        myStreamReader.Close();
                        myResponseStream.Close();
                        response.Close();
                        if (CompanyList.Count <= 0)
                        {
                            string pattner = @"addGoodsInfo\('(?<key>.*?)\|(?<value>.*?)'\)";
                            regex = new Regex(pattner);
                            var matchs = regex.Matches(outdata2);
                            CompanyList.Clear();
                            foreach (Match matchItem in matchs)
                            {
                                CompanyList.Add(matchItem.Groups["value"].Value.Trim(), matchItem.Groups["key"].Value.Trim());
                            }
                        }
                        #endregion
                    }
                }
                #endregion
                #region Post请求,设置所在地区
                url = string.Format("http://jy.ahyycg.cn:8080/Enterprise/RelationQuery/ProductCompanySubArea.aspx?PID={0}&returnUrl=/Enterprise/RelationQuery/RelationQueryUnpack.aspx", entity.产品流水号);
                request = (HttpWebRequest)WebRequest.Create(url);
                request.Headers["Pragma"] = "no-cache";
                request.Accept = "text/html, application/xhtml+xml, */*";
                request.AllowAutoRedirect = false;
                request.AutomaticDecompression = DecompressionMethods.GZip;
                request.Headers["Accept-Language"] = "zh-CH;en-US";
                request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
                request.KeepAlive = true;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                request.Timeout = 50000;
                request.Method = "POST";
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
                strAspNetPager1_pagesize = System.Configuration.ConfigurationManager.AppSettings["PageSize"];
                request.CookieContainer.Add(new Cookie("37", strAspNetPager1_pagesize, "/", "jy.ahyycg.cn"));
                request.Referer = url;
                StringBuilder postStrSbuilder1 = new StringBuilder();
                str__EVENTTARGET = "ctl00$ContentPlaceHolder1$ddlCity";
                str__EVENTARGUMENT = string.Empty;
                str__LASTFOCUS = string.Empty;
                str__VIEWSTATEGENERATOR = "4A4D21FA";
                var regex1 = new Regex(ViewStatePattner);
                var match1 = regex1.Match(outdata);
                if (match1 != null)
                    str__VIEWSTATE = match1.Groups["value"].Value;
                if (!this.CityList.ContainsKey(entity.所在地区.Trim()))
                    throw new Exception("未找到匹配的所在地区");
                strddlCity = this.CityList[entity.所在地区.Trim()];
                strddlhosname = string.Empty;
                strcompanyId = string.Empty;
                strAspNetPager1_input = "1";
                strhfdCompanyName = string.Empty;
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTTARGET"), HttpUtility.UrlEncode(str__EVENTTARGET));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTARGUMENT"), HttpUtility.UrlEncode(str__EVENTARGUMENT));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__LASTFOCUS"), HttpUtility.UrlEncode(str__LASTFOCUS));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATE"), HttpUtility.UrlEncode(str__VIEWSTATE));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATEGENERATOR"), HttpUtility.UrlEncode(str__VIEWSTATEGENERATOR));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$ddlCity"), HttpUtility.UrlEncode(strddlCity));//所在地区
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$hfdCompanyName"), HttpUtility.UrlEncode(strhfdCompanyName));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$ddlhosname"), HttpUtility.UrlEncode(strddlhosname));//医疗名称
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$companyId"), HttpUtility.UrlEncode(strcompanyId));//配送企业
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$AspNetPager1_input"), HttpUtility.UrlEncode(strAspNetPager1_input));
                postStrSbuilder1.AppendFormat("{0}={1}", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$AspNetPager1_pagesize"), HttpUtility.UrlEncode(strAspNetPager1_pagesize));
                byte[] postBytes1 = Encoding.UTF8.GetBytes(postStrSbuilder1.ToString());
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = postBytes1.Length;
                var postDataStream1 = request.GetRequestStream();
                postDataStream1.Write(postBytes1, 0, postBytes1.Length);
                postDataStream1.Close();
                postDataStream1.Dispose();
                response = (HttpWebResponse)request.GetResponse();
                UpdateLocalCookies(response.Cookies);
                myResponseStream = response.GetResponseStream();
                myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
                outdata4 = myStreamReader.ReadToEnd();
                myStreamReader.Close();
                myResponseStream.Close();
                response.Close();
                #endregion
            }
            else
            {
                #region Post请求,设置所在地区
                url = string.Format("http://jy.ahyycg.cn:8080/Enterprise/RelationQuery/ProductCompanySubArea.aspx?PID={0}&returnUrl=/Enterprise/RelationQuery/RelationQueryUnpack.aspx", entity.产品流水号);
                request = (HttpWebRequest)WebRequest.Create(url);
                request.Headers["Pragma"] = "no-cache";
                request.Accept = "text/html, application/xhtml+xml, */*";
                request.AllowAutoRedirect = false;
                request.AutomaticDecompression = DecompressionMethods.GZip;
                request.Headers["Accept-Language"] = "zh-CH;en-US";
                request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
                request.KeepAlive = true;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                request.Timeout = 50000;
                request.Method = "POST";
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
                strAspNetPager1_pagesize = System.Configuration.ConfigurationManager.AppSettings["PageSize"];
                request.CookieContainer.Add(new Cookie("37", strAspNetPager1_pagesize, "/", "jy.ahyycg.cn"));
                request.Referer = url;
                StringBuilder postStrSbuilder1 = new StringBuilder();
                str__EVENTTARGET = "ctl00$ContentPlaceHolder1$ddlCity";
                str__EVENTARGUMENT = string.Empty;
                str__LASTFOCUS = string.Empty;
                str__VIEWSTATEGENERATOR = "4A4D21FA";
                var regex1 = new Regex(ViewStatePattner);
                var match1 = regex1.Match(outdata);
                if (match1 != null)
                    str__VIEWSTATE = match1.Groups["value"].Value;
                if (!this.CityList.ContainsKey(entity.所在地区.Trim()))
                    throw new Exception("未找到匹配的所在地区");
                strddlCity = this.CityList[entity.所在地区.Trim()];
                strddlhosname = string.Empty;
                strcompanyId = string.Empty;
                strAspNetPager1_input = "1";
                strhfdCompanyName = string.Empty;
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTTARGET"), HttpUtility.UrlEncode(str__EVENTTARGET));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTARGUMENT"), HttpUtility.UrlEncode(str__EVENTARGUMENT));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__LASTFOCUS"), HttpUtility.UrlEncode(str__LASTFOCUS));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATE"), HttpUtility.UrlEncode(str__VIEWSTATE));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATEGENERATOR"), HttpUtility.UrlEncode(str__VIEWSTATEGENERATOR));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$ddlCity"), HttpUtility.UrlEncode(strddlCity));//所在地区
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$hfdCompanyName"), HttpUtility.UrlEncode(strhfdCompanyName));
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$ddlhosname"), HttpUtility.UrlEncode(strddlhosname));//医疗名称
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$companyId"), HttpUtility.UrlEncode(strcompanyId));//配送企业
                postStrSbuilder1.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$AspNetPager1_input"), HttpUtility.UrlEncode(strAspNetPager1_input));
                postStrSbuilder1.AppendFormat("{0}={1}", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$AspNetPager1_pagesize"), HttpUtility.UrlEncode(strAspNetPager1_pagesize));
                byte[] postBytes1 = Encoding.UTF8.GetBytes(postStrSbuilder1.ToString());
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = postBytes1.Length;
                var postDataStream1 = request.GetRequestStream();
                postDataStream1.Write(postBytes1, 0, postBytes1.Length);
                postDataStream1.Close();
                postDataStream1.Dispose();
                response = (HttpWebResponse)request.GetResponse();
                UpdateLocalCookies(response.Cookies);
                myResponseStream = response.GetResponseStream();
                myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
                outdata4 = myStreamReader.ReadToEnd();
                myStreamReader.Close();
                myResponseStream.Close();
                response.Close();
                #endregion
            }
            #region Post请求
            url = string.Format("http://jy.ahyycg.cn:8080/Enterprise/RelationQuery/ProductCompanySubArea.aspx?PID={0}&returnUrl=/Enterprise/RelationQuery/RelationQueryUnpack.aspx", entity.产品流水号);
            request = (HttpWebRequest)WebRequest.Create(url);
            request.Headers["Pragma"] = "no-cache";
            request.Accept = "text/html, application/xhtml+xml, */*";
            request.AllowAutoRedirect = false;
            request.AutomaticDecompression = DecompressionMethods.GZip;
            request.Headers["Accept-Language"] = "zh-CH;en-US";
            request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
            request.KeepAlive = true;
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
            request.Timeout = 50000;
            request.Method = "POST";
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(cookies);
            strAspNetPager1_pagesize = System.Configuration.ConfigurationManager.AppSettings["PageSize"];
            request.CookieContainer.Add(new Cookie("37", strAspNetPager1_pagesize, "/", "jy.ahyycg.cn"));
            request.Referer = url;
            StringBuilder postStrSbuilder2 = new StringBuilder();
            str__EVENTTARGET = string.Empty;
            str__EVENTARGUMENT = string.Empty;
            str__LASTFOCUS = string.Empty;
            str__VIEWSTATEGENERATOR = "4A4D21FA";
            var regex2 = new Regex(ViewStatePattner);
            var match2 = regex2.Match(outdata4);
            if (match2 != null)
                str__VIEWSTATE = match2.Groups["value"].Value;
            if (!this.CityList.ContainsKey(entity.所在地区.Trim()))
                throw new Exception("未找到匹配的所在地区");
            strddlCity = this.CityList[entity.所在地区.Trim()];
            if (!this.HospitalList.ContainsKey(entity.医疗名称.Trim()))
                throw new Exception("未找到匹配的医疗名称");
            strddlhosname = this.HospitalList[entity.医疗名称.Trim()];
            if (!this.CompanyList.ContainsKey(entity.配送企业.Trim()))
                throw new Exception("未找到匹配的配送企业");
            strcompanyId = this.CompanyList[entity.配送企业.Trim()];
            strAspNetPager1_input = "1";
            string strHzButton1 = "添加配送关系";
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTTARGET"), HttpUtility.UrlEncode(str__EVENTTARGET));
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__EVENTARGUMENT"), HttpUtility.UrlEncode(str__EVENTARGUMENT));
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__LASTFOCUS"), HttpUtility.UrlEncode(str__LASTFOCUS));
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATE"), HttpUtility.UrlEncode(str__VIEWSTATE));
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("__VIEWSTATEGENERATOR"), HttpUtility.UrlEncode(str__VIEWSTATEGENERATOR));
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$ddlCity"), HttpUtility.UrlEncode(strddlCity));//所在地区
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$hfdCompanyName"), HttpUtility.UrlEncode(strhfdCompanyName));
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$ddlhosname"), HttpUtility.UrlEncode(strddlhosname));//医疗名称
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$companyId"), HttpUtility.UrlEncode(strcompanyId));//配送企业
            postStrSbuilder2.AppendFormat("{0}={1}&", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$HzButton1"), HttpUtility.UrlEncode(strHzButton1));
            postStrSbuilder2.AppendFormat("{0}={1}", HttpUtility.UrlEncode("ctl00$ContentPlaceHolder1$AspNetPager1_pagesize"), HttpUtility.UrlEncode(strAspNetPager1_pagesize));
            byte[] postBytes2 = Encoding.UTF8.GetBytes(postStrSbuilder2.ToString());
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = postBytes2.Length;
            var postDataStream2 = request.GetRequestStream();
            postDataStream2.Write(postBytes2, 0, postBytes2.Length);
            postDataStream2.Close();
            postDataStream2.Dispose();
            response = (HttpWebResponse)request.GetResponse();
            UpdateLocalCookies(response.Cookies);
            myResponseStream = response.GetResponseStream();
            myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
            string outdata5 = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            response.Close();
            #endregion
            string resultPattner = "<script type=\"text/javascript\">alert\\('(?<value>.*?)'\\);</script></form>";
            var regexResult = new Regex(resultPattner);
            var matchResult = regexResult.Match(outdata5);
            if (matchResult != null)
            {
                return matchResult.Groups["value"].Value;
            }
            throw new Exception("未获取配送结果");
        }

        private void Complete(string LSH)
        {
            try
            {
                lock (lockobj)
                {
                    var rowCount = ExcelHelper.ExecuteCommand(CompleteSql,
                        new OleDbParameter[] {
                    new OleDbParameter("@是否添加", "是") ,
                    new OleDbParameter("@产品流水号", LSH) 
                });
                    if (rowCount <= 0)
                    {
                        throw new Exception(string.Format("处理流水号：{0},影响行数为0", LSH));
                    }
                }
            }
            catch (Exception ex)
            {
                OnError("数据库操作失败", ex.Message);
            }
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            var ds = ExcelHelper.GetReader(UnAddSql, new OleDbParameter[] { new OleDbParameter("@是否添加", "是") });
            var dt = ds.Tables[0];
            for (int i = 0; i < 5; i++)
            {
                UnComplete(dt.Rows[i]["产品流水号"].ToString());
            }
        }
        private void UnComplete(string LSH)
        {
            try
            {
                lock (lockobj)
                {
                    var rowCount = ExcelHelper.ExecuteCommand(CompleteSql,
                        new OleDbParameter[] {
                    new OleDbParameter("@是否添加", "否") ,
                    new OleDbParameter("@产品流水号", LSH) 
                });
                    if (rowCount <= 0)
                    {
                        throw new Exception(string.Format("处理流水号：{0},影响行数为0", LSH));
                    }
                }
            }
            catch (Exception ex)
            {
                OnError("数据库操作失败", ex.Message);
            }
        }

        private void GetCode()
        {
            try
            {
                string url = "http://jy.ahyycg.cn:8080/UserLogin.aspx";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Accept = "text/html, application/xhtml+xml, */*";
                request.AllowAutoRedirect = false;
                request.AutomaticDecompression = DecompressionMethods.GZip;
                request.Headers["Accept-Language"] = "zh-CH;en-US";
                request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
                request.KeepAlive = true;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                request.Timeout = 50000;
                request.Method = "GET";
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                //UpdateLocalCookies(response.Cookies);
                //response.Close();
                UpdateLocalCookies(response.Cookies);
                var myResponseStream = response.GetResponseStream();
                var myStreamReader = new StreamReader(myResponseStream, Encoding.UTF8);
                string outdata = myStreamReader.ReadToEnd();
                myStreamReader.Close();
                myResponseStream.Close();
                response.Close();


                url = "http://jy.ahyycg.cn:8080/CommonPage/Code.aspx";
                request = (HttpWebRequest)WebRequest.Create(url);
                request.Accept = "text/html, application/xhtml+xml, */*";
                request.AllowAutoRedirect = false;
                request.AutomaticDecompression = DecompressionMethods.GZip;
                request.Headers["Accept-Language"] = "zh-CH;en-US";
                request.Headers["Accept-Encoding"] = "gzip, deflate, sdch";
                request.KeepAlive = true;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                request.Timeout = 50000;
                request.Method = "GET";
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
                response = (HttpWebResponse)request.GetResponse();
                using (var stream = response.GetResponseStream())
                {
                    Image image = Image.FromStream(stream);
                    this.pictureBox1.Image = image;
                }
                UpdateLocalCookies(response.Cookies);
                response.Close();
            }
            catch (Exception ex)
            {
                OnError("获取验证码", ex.Message);
            }
        }

        public void UpdateLocalCookies(CookieCollection cookiesToUpdate)
        {
            if (cookiesToUpdate.Count == 0) return;

            if (cookies == null) { cookies = cookiesToUpdate; return; }

            foreach (Cookie toAdd in cookiesToUpdate)
            {
                bool found = false;
                if (cookies.Count > 0)
                {
                    foreach (Cookie originalCookie in cookies)
                    {
                        if (originalCookie.Name == toAdd.Name)
                        {
                            if (originalCookie.Domain == toAdd.Domain)
                            {
                                originalCookie.Value = toAdd.Value;
                                originalCookie.Domain = toAdd.Domain;
                                originalCookie.Expires = toAdd.Expires;
                                originalCookie.Version = toAdd.Version;
                                originalCookie.Path = toAdd.Path;
                                originalCookie.HttpOnly = toAdd.HttpOnly;
                                originalCookie.Secure = toAdd.Secure;
                                found = true;
                                break;
                            }
                        }
                    }
                }
                if (!found && toAdd.Domain != "")
                {
                    cookies.Add(toAdd);
                }
            }
        }

        private void txtCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (string.IsNullOrEmpty(this.txtCode.Text.Trim()))
                {
                    MessageBoxEx.Show(this, "验证码不能为空");
                    return;
                }
                this.btnLogin_Click(this.btnLogin, null);
            }
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            this.txtCode.Focus();
        }

        private void OnError(string Category, string ErrorMsg)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new MethodInvoker(() =>
                {
                    this.listBox1.Items.Add(string.Format("{0}:{1}", Category, ErrorMsg));
                }));
            }
            else
                this.listBox1.Items.Add(string.Format("{0}:{1}", Category, ErrorMsg));
            LogHelper.WriteError(Category, ErrorMsg);
        }
        private void SetEnable(bool IsEnable)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new MethodInvoker(() =>
                {
                    this.cbbUsers.Enabled = !IsEnable;
                    this.txtCode.ReadOnly = !IsEnable;
                    this.btnLogin.Enabled = IsEnable;
                    this.btnSetFile.Enabled = IsEnable;
                    this.btnSetPS.Enabled = IsEnable;
                }));
            }
            else
            {
                this.cbbUsers.Enabled = !IsEnable;
                this.txtCode.ReadOnly = !IsEnable;
                this.btnLogin.Enabled = IsEnable;
                this.btnSetFile.Enabled = IsEnable;
                this.btnSetPS.Enabled = IsEnable;
            }
        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            IsStop = true;
        }

        private void cbbUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedIndex = this.cbbUsers.SelectedIndex;
            if (selectedIndex == 1)
            {
                txtUserName.Text = "QS0239";
                txtPassword.Text = "635751";
            }
            else if (selectedIndex == 2)
            {
                txtUserName.Text = "QS0238";
                txtPassword.Text = "780785";
            }
            else
            {
                txtUserName.Text = string.Empty;
                txtPassword.Text = string.Empty;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            GetCode();
            this.txtCode.Text = string.Empty;
            this.txtCode.Focus();
        }
        private Cookie GetCookie()
        {
            int size = 300;
            StringBuilder SbCookie = new StringBuilder(size);
            if (API.InternetGetCookie("http://jy.ahyycg.cn", "ASP.NET_SessionId", SbCookie, ref size))
            {
                string result = SbCookie.ToString();
                var sessionid = result.Split(new char[] { '=' });
                var mycookie = new Cookie(sessionid[0], sessionid[1], "/", "jy.ahyycg.cn");
                return mycookie;
            }
            else
            {
                return null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //string baiduMainUrl = "http://www.baidu.com/";
            ////generate http request
            //HttpWebRequest req = (HttpWebRequest)WebRequest.Create(baiduMainUrl);

            ////add follow code to handle cookies
            //req.CookieContainer = new CookieContainer();
            //req.CookieContainer.Add(curCookies);

            //req.Method = "GET";
            ////use request to get response
            //HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            //txbGotBaiduid.Text = "";
            //foreach (Cookie ck in resp.Cookies)
            //{
            //    txbGotBaiduid.Text += "[" + ck.Name + "]=" + ck.Value;
            //    if (ck.Name == "BAIDUID")
            //    {
            //        gotCookieBaiduid = true;
            //    }
            //}

            //if (gotCookieBaiduid)
            //{
            //    //store cookies
            //    curCookies = resp.Cookies;
            //}
            //else
            //{
            //    MessageBox.Show("错误：没有找到cookie BAIDUID ！");
            //}


            //foreach (Cookie cookie in cookies)
            //{
            //    if (!API.InternetSetCookie("https://" + cookie.Domain.ToString() + ":8080", cookie.Name.ToString(), cookie.Value.ToString() + ";expires=Sun,22-Feb-2099 00:00:00 GMT"))
            //    {
            //        MessageBoxEx.Show(this, "失去连接请重新登陆");
            //    }
            //}
            //Process.Start("IExplore.exe", "http://jy.ahyycg.cn:8080/Default.aspx");
            int size = 300;
            StringBuilder SbCookie = new StringBuilder(size);
            if (API.InternetGetCookie("https://jy.ahyycg.cn", "ASP.NET_SessionId", SbCookie, ref size))
            {
                string result = SbCookie.ToString();
                var sessionid = result.Split(new char[] { '=' });
                var mycookie = new Cookie(sessionid[0], sessionid[1], "/", "jy.ahyycg.cn");
                if (API.InternetSetCookie("https://" + mycookie.Domain.ToString() + ":8080", mycookie.Name.ToString(), mycookie.Value.ToString() + ";expires=Sun,22-Feb-2099 00:00:00 GMT"))
                {
                    Process.Start("IExplore.exe", "http://jy.ahyycg.cn:8080/Default.aspx");
                }
            }
            else
            {
                string result = SbCookie.ToString();
                int errorCode = API.GetLastError();
            }
        }
    }
    public class Entity
    {
        public string 产品流水号 { get; set; }
        public string 所在地区 { get; set; }
        public string 医疗名称 { get; set; }
        public string 配送企业 { get; set; }
        public string 是否添加 { get; set; }
    }
    //if (!API.InternetSetCookie("http://jy.ahyycg.cn:8080", "ASP.NET_SessionId", cookie.ToString() + ";expires=Sun,22-Feb-2099 00:00:00 GMT"))
    //{
    //    MessageBox.Show(API.GetLastError().ToString());
    //}
    public class MyWorkThread
    {
        Thread thread;
        public MyWorkThread(Thread thread)
        {
            this.thread = thread;
        }

        public bool IsBackground
        {
            get
            {
                if (thread == null)
                    throw new Exception("线程为空");
                return false;
            }
            set
            {
                thread.IsBackground = value;
            }
        }
        public int WorkThreadId
        {
            get
            {
                if (thread == null)
                    throw new Exception("线程为空");
                return thread.ManagedThreadId;
            }
        }

        public void Start()
        {
            this.thread.Start();
        }
        private bool workComplete = false;
        public bool WorkComplete
        {
            get
            {
                return workComplete;
            }
            set
            {
                workComplete = value;
            }
        }
    }
}
