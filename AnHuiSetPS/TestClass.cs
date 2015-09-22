using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Text;

namespace AnHuiSetPS
{
    public class FullWebBrowserCookie
    {
        [System.Runtime.InteropServices.DllImport("wininet.dll", SetLastError = true)]
        public static extern bool InternetGetCookie(string url, string cookieName, StringBuilder cookieData, ref int size);

        private static int INTERNET_COOKIE_HTTPONLY = 0x00002000;
        [System.Runtime.InteropServices.DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetGetCookieEx(string pchURL, string pchCookieName,
            StringBuilder pchCookieData, ref System.UInt32 pcchCookieData,
            int dwFlags, IntPtr lpReserved);

        public static CookieContainer GetUriCookieContainer(Uri uri)
        {
            CookieContainer cookies = null;
            //定义Cookie数据的大小。
            int datasize = 256;
            StringBuilder cookieData = new StringBuilder(datasize);

            if (!InternetGetCookie(uri.ToString(), null, cookieData, ref datasize))
            {
                if (datasize < 0) return null;
                // 确信有足够大的空间来容纳Cookie数据。
                cookieData = new StringBuilder(datasize);
                if (!InternetGetCookie(uri.ToString(), null, cookieData, ref datasize)) return null;
            }
            try
            {
                string SessionId = GetSessionId(uri.ToString());
                if (SessionId.Trim().Length > 0)
                {
                    cookieData.Append(";" + SessionId);
                }
            }
            catch { }
            if (cookieData.Length > 0)
            {
                cookies = new CookieContainer();
                cookies.SetCookies(uri, cookieData.ToString().Replace(';', ','));
            }
            return cookies;
        }

        public static string GetSessionId(string url)
        {
            return GetCookie(url, "ASP.NET_SessionId");
        }
        public static string GetCookie(string url, string cookieName)
        {
            System.UInt32 size = 256;
            StringBuilder cookie = new StringBuilder(256);

            InternetGetCookieEx(url, cookieName, cookie, ref size,
                INTERNET_COOKIE_HTTPONLY, IntPtr.Zero);

            return cookie.ToString();
        }



        #region DLL Imports
        [SuppressUnmanagedCodeSecurity, SecurityCritical, DllImport("wininet.dll", EntryPoint = "InternetGetCookieExW", CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
        internal static extern bool InternetGetCookieEx([In] string Url, [In] string cookieName, [Out] StringBuilder cookieData, [In, Out] ref uint pchCookieData, uint flags, IntPtr reserved);
        #endregion
        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetGetCookieEx(string pchURL, string pchCookieName, StringBuilder pchCookieData, ref int pcchCookieData, int dwFlags, object lpReserved);
        public static string GetCookieString(string url)
        {
            // Determine the size of the cookie     
            int datasize = 256;
            StringBuilder cookieData = new StringBuilder(datasize);
            if (!InternetGetCookieEx(url, null, cookieData, ref datasize, 0x00002000, null))
            {
                if (datasize < 0)
                    return null;
                // Allocate stringbuilder large enough to hold the cookie     
                cookieData = new StringBuilder(datasize);
                if (!InternetGetCookieEx(url, null, cookieData, ref datasize, 0x00002000, null))
                    return null;
            }
            return cookieData.ToString();
        }
    }
}
