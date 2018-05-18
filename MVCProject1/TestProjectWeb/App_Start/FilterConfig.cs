using System.Web;
using System.Web.Mvc;
using TestProjectWeb.Utility;

namespace TestProjectWeb
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
            filters.Add(new CommonFilter());
        }
    }

    public class CommonFilter : ActionFilterAttribute {
        /// <summary>
        /// TODO
        /// </summary>
        /// <param name="filterContext">The filter context.</param>
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            HttpContextHelper.BuildContext(filterContext.HttpContext);
            base.OnActionExecuting(filterContext);
        }

        /// <summary>
        /// TODO
        /// </summary>
        /// <param name="filterContext"></param>
        public override void OnActionExecuted(ActionExecutedContext actionExecutedContext)
        {
            base.OnActionExecuted(actionExecutedContext);
        }

        /// <summary>
        /// TODO
        /// </summary>
        /// <param name="filterContext"></param>
        public override void OnResultExecuted(ResultExecutedContext resultExecutedContext)
        {
            HttpContextHelper.DestroyContext();
            base.OnResultExecuted(resultExecutedContext);
        }

    }
}