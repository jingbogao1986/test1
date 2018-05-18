using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(MVCApplication1.Startup))]
namespace MVCApplication1
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
