using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;

namespace TestProjectWeb.Utility
{
    class HttpContextHelper
    {
        private static readonly ThreadLocal<HttpContextHelper> _httpContext = new ThreadLocal<HttpContextHelper>();

        private readonly HttpRequestBase _httpRequestBase;

        private readonly HttpSessionStateBase _httpSessionStateBase;

        private HttpContextHelper(HttpContextBase httpContext)
        {
            _httpRequestBase = httpContext.Request;
            _httpSessionStateBase = httpContext.Session;
        }

        public static HttpContextHelper GetCurrent()
        {
            return _httpContext.Value;
        }

        public static void BuildContext(HttpContextBase httpContext)
        {
            var context = new HttpContextHelper(httpContext);
            _httpContext.Value = context;
        }

        public static void DestroyContext()
        {
            _httpContext.Value = null;
        }

        public string GetHttpRequestItem(string itemName)
        {
            return _httpRequestBase[itemName];
        }

        public string GetHttpRequestFile(string itemName)
        {
            return _httpRequestBase[itemName];
        }

        public object GetHttpSessionItem(string itemName)
        {
            return _httpSessionStateBase[itemName];
        }

        public void SetHttpSessionItem(string itemName, object itemValue)
        {
            _httpSessionStateBase[itemName] = itemValue;
        }

        public void AbandonHttpSession()
        {
            _httpSessionStateBase.Abandon();
        }

    }
}