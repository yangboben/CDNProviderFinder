using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel=Microsoft.Office.Interop.Excel;

namespace webspider
{
    class Program
    {
        static void Main(string[] args)
        {

            //string url = "www.gkkxd.com".Trim();     
            /*
            string html = tool.GetHtml(url);

           // System.IO.File.WriteAllText("test.txt", html);
            
            string[] urls,domains;

             urls = tool.GetUrlinHtml(url, html);
            //urls = tool.GetHrefinHtml(url, html);
             domains = urls.Select( m => tool.GetDomaininUrl(m)).ToArray();
          
            System.IO.File.WriteAllLines("url.txt",urls);
            System.IO.File.WriteAllLines("domains.txt", domains);


            string cdndomin = tool.GetDomainInHostName(tool.GetHostName(domains[0]));
            Console.Out.WriteLine(cdndomin);



            html = tool.GetHtml("http://www.beianbeian.com/piliang/chaxun?domains=" + cdndomin);


            HtmlDocument document = new HtmlDocument();
            document.LoadHtml(html);
            var t = document.DocumentNode.SelectSingleNode("//div[@id='show_table']").ChildNodes[1].ChildNodes[9].InnerText.Trim();
            
           
           // int n = document.DocumentNode.SelectSingleNode("//div[@id='show_table']").FirstChild.ChildNodes.Count;
        

            System.IO.File.WriteAllText("test.txt", html);
            
           Console.ReadKey();*/


            //tool.IfUseCDN("http://www.yovole.com/");


            tool.judgeFromExcel("D:/findcdn/CopyFY17List.xlsx", 15 ,16 ,18 ,1180);

        }




     
    }


    class tool
    {
        static string[] pic = {".jpg",".gif","ico"};
        static string[] video = {".wma",".mp4"};
        static string[] download = { ".js",".css",".html",".shtml",""};


       /// <summary>
       /// 
       /// </summary>
       /// <param name="filename">文件路径(最好是绝对路径)</param>
       /// <param name="col_domain">domain所在的列</param>
       /// <param name="col_icp">是否icp备案所在的列</param>
       /// <param name="col_result">结果所在的列</param>
       /// <param name="start">从第几行开始处理</param>
       /// <param name="end">到第几行结束,空着表示做到结尾</param>
        public static void judgeFromExcel(string filename,int col_domain, int col_icp, int col_result ,int start = 1,int end =-1) {
            
            
            
            Excel.Application table = new Excel.Application();
            Excel.Workbook workbook = table.Workbooks.Open(filename,0);
            Excel.Worksheet sheet = workbook.Worksheets.get_Item(1);
            Excel.Range range = sheet.UsedRange;
            
            if(start<2||start>range.Rows.Count){
                start = 2;
            }

            if(end <2||end>range.Rows.Count){
                end = range.Rows.Count;
            }

            for (int i = start; i <= end; i++)
            {

                string url = (range.Cells[col_domain][i] as Excel.Range).Value2;

                string result;
                if (!IfUseCDN(url, out result))
                    sheet.Cells[col_icp][i] = "未备案";
                else
                    sheet.Cells[col_icp][i] = null;

                sheet.Cells[col_result][i] = result;
                workbook.Save();
                Console.WriteLine("line"+i+" finished.");
            }
            workbook.Close();
            table = null;
            GC.Collect();
          

        }
        //if a website use cdn
        public static bool IfUseCDN(string domain,out string result)
        {
            if (domain == null)
            {
                result = "";
                return true;
            }

            bool icp=true;
     
            domain = domain.Trim();

            // get html and new a dictionary of resoucres
            Dictionary<string, cdntype> domains = new Dictionary<string, cdntype>();
            string html=""; 

             // test two kinds of protocol
            string protocol = GetProtocol(domain);
            if (protocol == "")
            {
                html = GetHtml("https://" + domain);
                if(html!="")
                {
                    domain = "https://" + domain;
                }
                else
                {
                    html = GetHtml("http://" + domain);
                    domain = "http://" + domain;
                }
            }
            else
            {

                html = GetHtml(domain);
            }
            


            // get all links in this html and all resoucres add to dictionary

            string[] links,urls;

            links = GetHrefinHtml(domain, html);
            urls = GetUrlinHtml(domain, html);

            adddomain(urls, domains, GetDomaininUrl(domain));
           
           // System.IO.File.WriteAllLines("url.txt", urls);
            //System.IO.File.WriteAllLines("domains.txt", links);
            // for all links get their resoucres and add them to dictionary
            
            foreach (string link in links)
            {
                
                try
                {
                    html = GetHtml(link);
                    urls = GetUrlinHtml(link,html);
                    adddomain(urls, domains, GetDomaininUrl(link)); 
                }
                catch(Exception e) {
                    html = null;
                    urls = null;
                }
            }
            
            // determine all the item in dictionary whether use cdn
            StringBuilder sb = new StringBuilder("");
            foreach (var item in domains)
            {
                try
                {
                    if (IsRedundancy(item.Key))
                        continue;
                    string HostName = GetHostName(item.Key);
                    string CDNdomain = GetDomainInHostName(HostName);       
                    string provider;
                    if (CDNdomain != GetDomainInHostName(GetDomaininUrl(domain))){ 
                        //备案了
                        if ( icp = beianchaxun(CDNdomain, out provider))                      
                            sb.AppendLine(item.Key + "  是否使用图片加速:" + item.Value.pic + "  是否使用视频加速:" + item.Value.video + "  是否使用下载加速:" + item.Value.download + "   provider:" + provider);                        
                        else
                            sb.AppendLine(item.Key + "  是否使用图片加速:" + item.Value.pic + "  是否使用视频加速:" + item.Value.video + "  是否使用下载加速:" + item.Value.download + "   provider:" + CDNdomain);
                    }
                }
                catch { }
            }
            if (sb.ToString() == "")
                sb.Append("没有发现使用cdn");
            // outpu result
            System.IO.File.WriteAllText("result.txt", sb.ToString());

            result = sb.ToString();

            return icp;
            
        }

        /// <summary>
        /// get html of this url
        /// </summary>
        /// <param name="Url"> the url you want</param>
        /// <returns> html as string </returns>
        public static string GetHtml(string Url)
        {

            try
            {
                HttpWebRequest wReq = (HttpWebRequest)WebRequest.Create(Url);
                wReq.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.1.4322; .NET CLR 2.0.50215; CrazyCoder.cn;www.aligong.com)";

                wReq.Referer = Url;

                HttpWebResponse wResp = wReq.GetResponse() as HttpWebResponse;
                System.IO.Stream respStream = wResp.GetResponseStream();
   
                System.IO.StreamReader reader = new System.IO.StreamReader(respStream, Encoding.UTF8);
                string content = reader.ReadToEnd();

                reader.Close();
                reader.Dispose();
                return content;
            } 
            catch (Exception e) { }
           
            return "";
        }

        /// <summary>
        /// get all resource in the url and html
        /// </summary>
        /// <param name="url"> the target url</param>
        /// <param name="html"> the target html</param>
        /// <returns> all resources url in the html </returns>
        public static string[] GetUrlinHtml(string url, string html)
        {                    
            string strRef = @"(?:src|SRC|href|HREF)[ ]*=[ ]*[""']([^""'#>]+)[""']";
            var matches = new Regex(strRef).Matches(html);

            return matches.Cast<Match>().Select(m => m.Groups[1].Value)
                                        .Select(m => m.StartsWith("//") ? GetProtocol(url) + ":" + m : m)                                  
                                        //.Select(m => m.StartsWith("http")? m : url+"/"+m)
                                        .ToArray();
        }


        /// <summary>
        /// 只获取当前网页domain下的 link
        /// </summary>
        /// <param name="url"> 当前网站的 url</param>
        /// <param name="html"> 当前得到的页面</param>
        /// <returns>所有link的数组</returns>

        public static string[] GetHrefinHtml(string url, string html)
        {
            if (url.EndsWith("/"))
                url = url.Substring(0, url.Length - 1);
            string strRef = @"(?:href|HREF|action|ACtion|Action)[ ]*=[ ]*[""']([^""'#>]+)[""']";
            var matches = new Regex(strRef).Matches(html);            
            return matches.Cast<Match>().Select(m => m.Groups[1].Value)
                                        .Where(m => m.StartsWith("/")&!m.StartsWith("//"))
                                        .Select(m => url + m)
                                     // .Select(m => m.StartsWith("//")?GetProtocol(url)+":"+m:m)                                     
                                      //.Select(m=> m.StartsWith("http")?m:url+"/"+m)  
                                        .Where(m=> !m.EndsWith(".css")&!m.EndsWith(".js"))
                                        .ToArray();
        }

        /// <summary>
        /// determine what protocol the web use 
        /// </summary>
        /// <param name="url"> the url the website</param>
        /// <returns></returns>
        public static string GetProtocol(string url){
            if (url == null)
                return "";
              if(url.StartsWith("https"))
                  return "https";
              if(url.StartsWith("http"))
                  return "http";
              return "";
        }


        /// <summary>
        /// get the domain the resource belongs to
        /// </summary>
        /// <param name="url"> the url gets from html getUrlinHtml() </param>
        /// <returns>如果是当前网站内部的，返回的是"" others like wow.178.com </returns>
        public static string GetDomaininUrl(string url)
        {
            string strRef = @"(?:http://|https://)([^""'#>/]+)/?";

            Match match = new Regex(strRef).Match(url);
            return match.Groups[1].Value;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="url">the resources want to judge</param>
        /// <returns>-1 表示不是任何类型的，1 表示图片，2 表示视频 ，3表示下载</returns>
        public static int GetType(string url) {
              url = url.Trim();

            //judge if it is a file
              string isfile = @".\w+$";
            if (!new Regex(isfile).IsMatch(url))
                  return -1;

            //judge if it is a pic
            foreach (string i in pic) {
                if (url.EndsWith(i))
                    return 1;
            }

            //judge if it is a video
            foreach (string i in video)
            {
                if (url.EndsWith(i))
                    return 2;
            }

            //judge if it is a download
            foreach (string i in download)
            {
                if (url.EndsWith(i))
                    return -1;
            }

            return 3;
        }


        /// <summary>
        ///  nslookup input is the domain dealed
        /// </summary>
        /// <param name="domain">domain dealed like www.baidu.com</param>
        /// <returns>domain with cdn like opthw.xdwscache.speedcdns.com</returns>

        public static string GetHostName(string domain) {
            
            IPHostEntry hostinfo = Dns.Resolve(domain);
            return hostinfo.HostName;
        }
        /// <summary>
        /// get domain to icp inquire
        /// </summary>
        /// <param name="HostName"> the hostname nslookup get like opthw.xdwscache.speedcdns.com</param>
        /// <returns> hostname can use to icp inquirment  like speedcdns.com </returns>
        public static string GetDomainInHostName(string HostName) {
            string strRef = @"[^""'#>/.]?.([^""'#>/.]+(?:.[^""'#>/.]+|.com.cn))$";

            Match match = new Regex(strRef).Match(HostName);
            return match.Groups[1].Value;
         
        }

        /// <summary>
        ///   beianchaxun
        /// </summary>
        /// <param name="domain"> the domain you want to inquire (by GetDomainInHostName())</param>
        /// <returns></returns>
        public static bool beianchaxun(string domain, out string provider) {

            string html = GetHtml("http://www.beianbeian.com/piliang/chaxun?domains=" + domain);
           
            HtmlDocument document = new HtmlDocument();
            document.LoadHtml(html);
            string t = document.DocumentNode.SelectSingleNode("//div[@id='show_table']").ChildNodes[1].ChildNodes[9].InnerText.Trim();
       
            if (t == "未备案")
            {
                provider = null;
                return false;
            }
            else
            {
                provider = document.DocumentNode.SelectSingleNode("//div[@id='show_table']").ChildNodes[1].ChildNodes[9].ChildNodes[9].InnerText;
                return true;
            }
        }


        /// <summary>
        /// 判断是否是一像政府网站图标的垃圾信息
        /// </summary>
        /// <param name="domain"> 要判断的域名</param>
        /// <returns></returns>
        public static bool IsRedundancy(string domain)
        {
              string strRef = @".gov[.$]";
              if (new Regex(strRef).IsMatch(domain))
                  return true;



              return false;
        }
        public static void adddomain(string[] urls, Dictionary<string, cdntype> domains,string domain)
        {
            foreach (string url in urls)
            {
                int type = GetType(url);
                if (type < 0)
                    continue;
                string url_domain = GetDomaininUrl(url);
                if (url_domain == "")
                    url_domain = domain;
                
                if (!domains.ContainsKey(url_domain))
                    domains[url_domain] = new cdntype();

                //向其中添加信息
                switch (type)
                {
                    case 1: domains[url_domain].pic = true;
                        break;
                    case 2: domains[url_domain].video = true;
                        break;
                    case 3: domains[url_domain].download = true;
                        break;
                    default:
                        break;
                }
            }
        }
    }


    class cdntype {
        public bool pic {get; set;}
        public bool video { get; set; }
        public bool download { get; set; }
    }
}
