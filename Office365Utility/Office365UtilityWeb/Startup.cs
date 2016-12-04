using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Office365UtilityWeb.Startup))]
namespace Office365UtilityWeb
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
