using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TestProjectWeb.Utility;

namespace TestProjectWeb.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Calculate()
        {
            string result = string.Empty;

            try
            {
                // 获取计算参数
                var quantity = HttpContextHelper.GetCurrent().GetHttpRequestItem("quantity");
                var level = HttpContextHelper.GetCurrent().GetHttpRequestItem("level");
                var solution = HttpContextHelper.GetCurrent().GetHttpRequestItem("solution");
                var quality = HttpContextHelper.GetCurrent().GetHttpRequestItem("quality");

                // 定义返回对象
                object objRtn = new object();

                // 获取源文件路径。 TBD: 变成系统资源。
                string filePath = Server.MapPath("~/Resources/抽样方案.xlsm");

                ExcelMacroUtil helper = new ExcelMacroUtil();

                // 根据用户页面输入，调用Excel的宏，设置输入字段单元格
                helper.RunExcelMacro(filePath, "SetValue", new object[] { quantity, level, solution, quality }, out objRtn, false);

                // 调用计算的宏，拿到返回值
                helper.RunExcelMacro(filePath, "GetResult", new object[] { }, out objRtn, false);

                result = (string)objRtn;
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }

            return Json(result);
        }

        #region Helper
        public static string GetJsonString(object obj)
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(obj);
        }

        public static void HandleException(Exception ex)
        {
            //  TBD
            throw (ex);
        }
        #endregion
    }
}
