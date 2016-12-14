using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(GMailSync.Startup))]
namespace GMailSync
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
