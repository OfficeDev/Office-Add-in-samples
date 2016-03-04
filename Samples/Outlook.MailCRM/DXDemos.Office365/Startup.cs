using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(DXDemos.Office365.Startup))]
namespace DXDemos.Office365
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
            app.MapSignalR();
        }
    }
}
