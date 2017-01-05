using PasswordUtility;
using PasswordUtility.PasswordGenerator;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

//網站共用功能函式庫
namespace Common.tools {

    /// <summary>
    /// 網站共用的功能
    /// </summary>
    public class Utility {

        #region *** 判斷字串的內容是否為數值 ***

        /// <summary>
        /// 判斷字串的內容是否為數值 (可判斷整數，亦可判斷帶有小數點的浮點數)
        /// </summary>
        /// <param name="s">要作判斷的字串</param>
        /// <returns>bool</returns>
        public static bool isStringContentNumeric(string s) {
            if (s != null) {
                int n;
                double d;
                if (int.TryParse(s.Trim(), out n)) {
                    return true;    //是數值 (整數)，包含：0、負數
                } else if (double.TryParse(s.Trim(), out d)) {
                    return true;    //是數值 (浮點數)，包含：0.0、負數
                } else {
                    return false;   //不是數值
                }
            } else {
                return false;       //不是數值
            }
        }

        /// <summary>
        /// 判斷字串的內容是否為整數 (只能判斷整數，無法判斷帶有小數點的浮點數)
        /// </summary>
        /// <param name="s">要作判斷的字串</param>
        /// <returns>bool</returns>
        public static bool isStringContentInteger(string s) {
            if (s != null) {
                int n;
                if (int.TryParse(s.Trim(), out n)) {
                    return true;    //是數值 (整數)，包含：0、負數
                } else {
                    return false;   //不是數值
                }
            } else {
                return false;       //不是數值
            }
        }

        #endregion

        #region *** 取得 Server IP 或使用者 IP ***

        /// <summary>
        /// 取得 Server IP (IPv4)
        /// 從本機執行會得到本機的 IP，放在 Server 上會得到 Server 的 IP 。
        /// </summary>
        /// <returns>string</returns>
        public static string getServerIP() {
            string IP4Address = string.Empty;

            foreach (IPAddress IPA in Dns.GetHostAddresses(Dns.GetHostName())) {
                if (IPA.AddressFamily.ToString().ToUpperInvariant().Equals("INTERNETWORK")) //InterNetwork 或 InterNetworkV6
                {
                    IP4Address = IPA.ToString().Trim();
                    break;
                }
            }

            //從本機執行會得到本機的 IP，放在 Server 上會得到 Server 的 IP
            return IP4Address;
        }

        /// <summary>
        /// 取得使用者 IP (IPv4)
        /// 從本機執行會得到「::1」字樣，放在 Server 上會得到 user 的 IP 。
        /// </summary>
        /// <returns>string</returns>
        public static string getUserIP() {
            string strUserIP = "";

            //做法1:
            strUserIP = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"].ToString().Trim(); //::1
            //Response.Write(Request.ServerVariables["HTTP_VIA"].ToString() + "<br />");    //error:並未將物件參考設定為物件的執行個體
            //Response.Write(Request.ServerVariables["HTTP_X_FORWARDED_FOR"].ToString());   //error:並未將物件參考設定為物件的執行個體
            //Response.Write(Page.Request.UserHostAddress + "<br />");                      //::1
            //Response.Write(HttpContext.Current.Request.UserHostAddress + "<br />");       //::1

            //做法2:
            //System.Web.HttpContext current = System.Web.HttpContext.Current;
            //string RemoteIP = "";
            //string innerIP = current.Request.Form["innerIP"] != null ? current.Request.Form["innerIP"].ToString() : string.Empty;
            //if (!string.IsNullOrEmpty(innerIP))
            //    RemoteIP = innerIP;
            //else
            //{
            //    string HTTP_VIA = current.Request.ServerVariables["Remote_Addr"] != null ? current.Request.ServerVariables["Remote_Addr"].ToString() : string.Empty;
            //    string HTTP_X_FORWARDED_FOR = string.Empty;
            //    if (!string.IsNullOrEmpty(HTTP_VIA) && HTTP_VIA.Contains("ebc.net.tw"))
            //        HTTP_X_FORWARDED_FOR = current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"] != null ? current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"].ToString() : string.Empty;

            //    if (!string.IsNullOrEmpty(HTTP_X_FORWARDED_FOR))
            //        RemoteIP = HTTP_X_FORWARDED_FOR;
            //    else
            //        RemoteIP = current.Request.ServerVariables["Remote_Addr"].ToString();
            //}

            //return RemoteIP;

            //從本機執行會得到「::1」字樣，放在 Server 上會得到 user 的 IP
            return strUserIP;

        }

        #endregion

        #region *** 產生不重複的亂數 ***

        /// <summary> 
        /// 產生不重複的亂數 (「產生亂數的範圍上限」至「「產生亂數的範圍下限」」之間的個數，不可小於「產生亂數的數量」，否則會跑不出迴圈)
        /// </summary> 
        /// <param name="intLower"></param>產生亂數的範圍下限 
        /// <param name="intUpper"></param>產生亂數的範圍上限 
        /// <param name="intNum"></param>產生亂數的數量 
        /// <returns>List<int></returns> 
        public static List<int> makeRand(int intLower, int intUpper, int intNum) {
            List<int> arrayRand = new List<int>();

            Random random = new Random((int)DateTime.Now.Ticks);
            int intRnd;
            while (arrayRand.Count < intNum) {
                intRnd = random.Next(intLower, intUpper + 1);
                if (!arrayRand.Contains(intRnd)) {
                    arrayRand.Add(intRnd);
                }
            }

            return arrayRand;
        }

        //引用此函數的方式:
        //System.Collections.Generic.List<int> rnd = makeRand(1, 8, 5);
        //foreach(int i in rnd.ToArray())
        //{
        //    Response.Write(i + ", ");
        //}

        #endregion

        #region *** 取得文章多久前發布的距離時間 ***

        /// <summary> 
        /// 判斷文章多久前發布的距離時間
        /// </summary> 
        /// <param name="startDate">文章發佈時間</param>
        /// <returns>string</returns> 
        public static string getStringDateRegin(string startDate) {
            DateTime STime = DateTime.Parse(startDate); //起始日
            DateTime ETime = Convert.ToDateTime(DateTime.Now);//結束日
            TimeSpan Total = ETime.Subtract(STime); //日期相減
            int mDays = Total.Days;
            string Rdate = "剛剛";
            if (mDays < 1) {
                int TotalHours = (int)Total.TotalHours;
                if (TotalHours < 1) {
                    int Minutes = Total.Minutes;
                    if (Minutes > 10) {
                        Rdate = Minutes + "分前"; //共幾分
                    }
                } else {
                    Rdate = TotalHours + "小時前"; //總共多少小時 
                }

            } else {
                Rdate = mDays + "天前"; //共幾天 
            }

            return Rdate;
        }

        #endregion

        #region *** 依傳入的日期時間，取得與系統時間的差距(小時) ***

        /// <summary>
        /// 依傳入的日期時間，取得與系統時間的差距(小時)
        /// </summary>
        /// <param name="DateAndTime">傳入的日期時間</param>
        /// <returns>int</returns>
        public static int getTimeSpanByHour(string DateAndTime) {
            int intReturnHour = 0;
            TimeSpan ts;

            try {
                ts = DateTime.Now - Convert.ToDateTime(DateAndTime);
                intReturnHour = Convert.ToInt32(ts.TotalHours);
            } catch (Exception) {
                throw;
            }

            return intReturnHour;
        }

        #endregion

        #region *** 轉址至路由定義的key名 ***

        /// <summary>
        /// 轉址至路由定義的key名
        /// </summary>
        /// <param name="routeKey">路由定義的key名</param>
        public static void toRoute(string routeKey) {
            HttpContext.Current.Response.RedirectToRoute(routeKey);
        }

        #endregion

        #region *** 將字串轉換為Base64字串 ***

        /// <summary>
        /// 將字串轉換為Base64字串
        /// </summary>
        /// <param name="originStr">來源字串</param>
        /// <param name="encoding">編碼方式</param>
        /// <returns>string</returns>
        public static string stringToBase64(string originStr, Encoding encoding) {
            string outPut = null;
            if (!string.IsNullOrEmpty(originStr)) {
                outPut = Convert.ToBase64String(encoding.GetBytes(originStr));
            }

            return outPut;
        }

        #endregion

        #region *** 利用正則表示式判斷字元是否為中文字，返回指定長度的字串 ***

        /// <summary>
        /// 利用正則表示式判斷字元是否為中文字，返回指定長度的字串
        /// </summary>
        /// <param name="text">要檢測的字串</param>
        /// <param name="limit">限制要截取的長度</param>
        /// <param name="outCount">回傳計算過後的字元數量</param>
        /// <returns>string</returns>
        public static string subUnicodeCharForLimit(string text, int limit, out int outCount) {
            StringBuilder sb = new StringBuilder();
            char[] cAry = text.ToCharArray();
            string pattern = "[^\x00-\x7F]+";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            int count = 0;
            foreach (char c in cAry) {
                if (count >= limit) { break; }
                if (regex.IsMatch(c.ToString())) {      //中文字
                    count += 2;
                } else {                                //非中文字
                    count++;
                }
                sb.Append(c.ToString());
            }

            outCount = count;
            return sb.ToString();
        }

        #endregion

        #region *** 檢查現在時間是否介於指定時間區間 ***

        /// <summary>
        /// 檢查現在時間是否介於指定時間區間
        /// </summary>
        /// <param name="beginDT">要檢查的開始時間</param>
        /// <param name="endDT">要檢查的結束時間</param>
        /// <returns>bool</returns>
        public static bool checkNowDateTimeInScope(DateTime beginDT, DateTime endDT) {
            DateTime now = DateTime.Now;
            return (beginDT <= now && now <= endDT);
        }

        #endregion

        #region *** 改變圖檔尺寸 ***

        /// <summary>
        /// 取得圖檔串流並改變圖檔像素比與解析度
        /// </summary>
        /// <param name="stream">圖片串流</param>
        /// <param name="widthPx">要改變的圖片寬度像素值</param>
        /// <param name="heightPx">要改變的圖片高度像素值</param>
        /// <param name="xDpi">要改變的圖片橫向解析度(非必填)</param>
        /// <param name="yDpi">要改變的圖片縱向解析度(非必填)</param>
        /// <returns>Bitmap</returns>
        public static Bitmap changeImageSize(Stream stream, int widthPx, int heightPx, float xDpi = 0, float yDpi = 0) {
            Bitmap ResultNewBitMap = null;

            try {
                Bitmap bitMap = (Bitmap)Image.FromStream(stream, true, true);
                ResultNewBitMap = new Bitmap(bitMap, widthPx, heightPx);
                ResultNewBitMap.SetResolution(((xDpi.Equals(0)) ? bitMap.HorizontalResolution : xDpi),
                    ((yDpi.Equals(0)) ? bitMap.VerticalResolution : yDpi));
            } catch (Exception) {
                throw;
            }

            return ResultNewBitMap;
        }

        /// <summary>
        /// 取得實體圖檔並改變圖檔像素比與解析度
        /// </summary>
        /// <param name="filePath">實體檔案路徑</param>
        /// <param name="widthPx">要改變的圖片寬度像素值</param>
        /// <param name="heightPx">要改變的圖片高度像素值</param>
        /// <param name="xDpi">要改變的圖片橫向解析度(非必填)</param>
        /// <param name="yDpi">要改變的圖片縱向解析度(非必填)</param>
        /// <returns>Bitmap</returns>
        public static Bitmap changeImageSize(string filePath, int widthPx, int heightPx, float xDpi = 0, float yDpi = 0) {
            Bitmap ResultNewBitMap = null;

            try {
                Bitmap bitMap = (Bitmap)Image.FromFile(filePath, true);
                ResultNewBitMap = new Bitmap(bitMap, widthPx, heightPx);
                ResultNewBitMap.SetResolution(((xDpi.Equals(0)) ? bitMap.HorizontalResolution : xDpi),
                    ((yDpi.Equals(0)) ? bitMap.VerticalResolution : yDpi));
            } catch (Exception) {
                throw;
            }

            return ResultNewBitMap;
        }

        #endregion

        #region *** 取得 web.config 或 app.config AppSettings 值 ***

        /// <summary>
        /// 取得 web.config 或 app.config AppSettings 值
        /// </summary>
        /// <param name="key">AppSettings key 名</param>
        /// <returns>string</returns>
        public static string getAppSettings(string key) {
            return ConfigurationManager.AppSettings[key];
        }

        /// <summary>
        /// 取得 web.config 或 app.config AppSettings 值
        /// </summary>
        /// <param name="key">AppSettings key 名</param>
        /// <param name="defaultValue">若無取得相對應資料，預設傳回的值</param>
        /// <returns>string</returns>
        public static string getAppSettings(string key, string defaultValue) {
            string result = ConfigurationManager.AppSettings[key];
            return ((string.IsNullOrEmpty(result)) ? defaultValue : result);
        }

        #endregion

        #region *** 取得網頁第一次輸出的靜態內容 ***

        /// <summary>
        /// 取得網頁內容字串
        /// </summary>
        /// <param name="url">網頁 url</param>
        /// <returns>string</returns>
        public static string getWebContent(string url) {
            return getWebContent(url, Encoding.Default);
        }

        /// <summary>
        /// 取得網頁內容字串
        /// </summary>
        /// <param name="url">網頁 url</param>
        /// <param name="encoding">指定的編碼方式</param>
        /// <exception cref="Exception"></exception>
        /// <returns>List<object></returns>
        public static string getWebContent(string url, Encoding encoding) {
            string result = string.Empty;
            using (WebClient webClient = new WebClient()) {
                webClient.Proxy = null;
                webClient.Encoding = encoding;

                try {
                    result = webClient.DownloadString(url);
                } catch (Exception) {
                    throw;
                }
            }

            return result;
        }

        #endregion

        #region *** Exception 例外資訊處理 ***

        /// <summary>
        /// 取得例外資訊行號
        /// </summary>
        /// <param name="ex">Excetion例外物件</param>
        /// <returns>int</returns>
        public static int getExLineNumber(Exception ex) {
            return new StackTrace(ex, true).GetFrame(0).GetFileLineNumber();
        }

        /// <summary>
        /// 取得例外資訊欄位號
        /// </summary>
        /// <param name="ex">Excetion例外物件</param>
        /// <returns>int</returns>
        public static int getExColumnNumber(Exception ex) {
            return new StackTrace(ex, true).GetFrame(0).GetFileColumnNumber();
        }

        #endregion

        #region *** 文字檔案讀寫 ***

        /// <summary>
        /// 建立或加入文字檔案內容
        /// </summary>
        /// <param name="dir">文字檔案存放路徑</param>
        /// <param name="fileNameWithExt">檔名 + 副檔名</param>
        /// <param name="contentList">文字檔案內容集合</param>
        /// <returns>List<object></returns>
        public static List<object> createOrAppendFile(string dir, string fileNameWithExt, List<string> contentList) {
            List<object> resultList = new List<object>();

            if (!Directory.Exists(dir)) {
                Directory.CreateDirectory(dir);
            }

            try {
                File.AppendAllLines(string.Concat(dir, fileNameWithExt), contentList, Encoding.UTF8);
                resultList.Add(1);
            } catch (Exception ex) {
                resultList.Add(0);
                resultList.Add(ex.Message);
            }

            return resultList;
        }

        #endregion

        #region *** SqlDataReader 功能擴充 ***

        /// <summary>
        /// SqlDataReader 取得 null 值時轉換成指定預設值
        /// </summary>
        /// <param name="reader">SqlDataReader</param>
        /// <param name="readerIdx">SqlDataReader 索引值</param>
        /// <param name="defaultValue">預設值</param>
        /// <returns>object</returns>
        public static object readerSafeValue(SqlDataReader reader, int readerIdx, object defaultValue) {
            if (reader.IsDBNull(readerIdx)) {
                return defaultValue;
            } else {
                return reader[readerIdx];
            }
        }

        #endregion

        #region *** Cookie 操作 ***

        /// <summary>
        /// 寫入 cookie 多值集合 (處理成功回傳空字串，失敗回傳例外訊息)
        /// </summary>
        /// <param name="cookieName">cookie 名稱</param>
        /// <param name="cookieValues">cookie 鍵值集合</param>
        /// <param name="expires">cookie 到期日期</param>
        /// <param name="isHttpOnly">cookie 是否允許 client 端存取 (預設可以)</param>
        /// <returns>Tuple<bool, string></returns>
        public static string setCookieValues(string cookieName, Dictionary<string, object> cookieValues,
            DateTime expires, bool isHttpOnly = false) {

            return setCookieValues(cookieName, (object)cookieValues, expires, isHttpOnly);
        }

        /// <summary>
        /// 寫入 cookie 單一值 (處理成功回傳空字串，失敗回傳例外訊息)
        /// </summary>
        /// <param name="cookieName">cookie 集合名稱</param>
        /// <param name="cookieValues">cookie 值</param>
        /// <param name="expires">cookie 到期日期</param>
        /// <param name="isHttpOnly">cookie 是否允許 client 端存取 (預設可以)</param>
        /// <returns>string</returns>
        public static string setCookieValues(string cookieName, object cookieValues,
            DateTime expires, bool isHttpOnly = false) {

            string result = string.Empty;
            HttpCookie cookie;

            if (HttpContext.Current.Request.Cookies[cookieName] == null) {
                cookie = new HttpCookie(cookieName);
            } else {
                cookie = HttpContext.Current.Request.Cookies.Get(cookieName);
            }

            if (cookieValues is Dictionary<string, object>) {
                foreach (var item in cookieValues as Dictionary<string, object>) {
                    cookie[item.Key] = Convert.ToString(item.Value);
                }
            } else {
                cookie.Value = Convert.ToString(cookieValues);
            }

            cookie.Expires = expires;
            cookie.HttpOnly = isHttpOnly;

            try {
                HttpContext.Current.Response.SetCookie(cookie);
            } catch (Exception ex) {
                result = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 取得 cookie 單一值
        /// </summary>
        /// <param name="cookieName">cookie 名稱</param>
        /// <param name="keyName">cookie 集合鍵名</param>
        /// <returns>string</returns>
        public static string getCookieValue(string cookieName, string keyName = null) {
            string result = string.Empty;

            if (HttpContext.Current.Request.Cookies[cookieName] != null) {
                if (keyName == null) {
                    result = HttpContext.Current.Request.Cookies.Get(cookieName).Value;
                } else {
                    result = HttpContext.Current.Request.Cookies.Get(cookieName).Values[keyName];
                }
            }

            return result;
        }

        /// <summary>
        /// 取得 cookie 集合內容
        /// </summary>
        /// <param name="cookieName">cookie 集合名稱</param>
        /// <returns></returns>
        public static NameValueCollection getCookieValueCollection(string cookieName) {
            NameValueCollection result = new NameValueCollection();

            if (HttpContext.Current.Request.Cookies[cookieName] != null) {
                result = HttpContext.Current.Request.Cookies.Get(cookieName).Values;
            }

            return result;
        }

        /// <summary>
        /// 完整刪除 cookie (處理成功回傳空字串，失敗回傳例外訊息)
        /// </summary>
        /// <param name="cookieName">cookie 名稱</param>
        /// <returns>string</returns>
        public static string deleteCookie(string cookieName) {
            string result = string.Empty;

            if (HttpContext.Current.Request.Cookies[cookieName] != null) {
                HttpCookie cookie = HttpContext.Current.Request.Cookies.Get(cookieName);
                cookie.Expires = DateTime.Now.AddDays(-1d);

                try {
                    HttpContext.Current.Response.SetCookie(cookie);
                } catch (Exception ex) {
                    result = ex.Message;
                }
            }

            return result;
        }

        #endregion

        #region *** 密碼規則建立與檢查 ***

        public enum PasswordRule { NORMAL, CHECK_UPPERCASE, CHECK_UPPERCASE_WITH_NUMERIC };

        /// <summary>
        /// 依據指定的規則產生密碼
        /// </summary>
        /// <param name="length">密碼長度</param>
        /// <param name="useUpperCase">是否包含大寫字元</param>
        /// <param name="useNumeric">是否包含數字字元</param>
        /// <param name="useSpecialChar">是否包含特殊字元</param>
        /// <returns>string</returns>
        public static string generatePassword(int length, bool useUpperCase, bool useNumeric, bool useSpecialChar = false) {
            return PwGenerator.Generate(length, useUpperCase, useNumeric, useSpecialChar).ReadString();
        }

        /// <summary>
        /// 檢查密碼規則
        /// </summary>
        /// <param name="password">要檢查的密碼字串</param>
        /// <param name="minLength">最小密碼長度</param>
        /// <param name="maxLength">最大密碼長度(預設不限制)</param>
        /// <param name="rule">密碼檢查規則列舉</param>
        /// <returns>bool</returns>
        public static bool checkPasswordRule(string password, int minLength, int maxLength = 0, PasswordRule rule = PasswordRule.NORMAL) {
            string pattern = string.Empty;
            switch (rule) {
                case PasswordRule.NORMAL:
                    pattern = @"^(?=.*[a-z]).{{0},{1}}$";
                    break;
                case PasswordRule.CHECK_UPPERCASE:
                    pattern = @"^(?=.*[a-z])(?=.*[A-Z]).{{0},{1}}$";
                    break;
                case PasswordRule.CHECK_UPPERCASE_WITH_NUMERIC:
                    pattern = @"^(?=.*\d)(?=.*[a-z])(?=.*[A-Z]).{{0},{1}}$";
                    break;
            }

            pattern = string.Format(pattern, minLength, ((maxLength.Equals(0)) ? string.Empty : maxLength.ToString()));
            return new Regex(pattern).IsMatch(password);
        }

        /// <summary>
        /// 檢查密碼強度
        /// </summary>
        /// <param name="password">要檢查的密碼字串</param>
        /// <returns>uint</returns>
        public static uint checkPasswordStrenfth(string password) {
            return QualityEstimation.EstimatePasswordBits(password.ToCharArray());
        }

        #endregion

    }
}
