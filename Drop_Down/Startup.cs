using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Drop_Down.Startup))]
namespace Drop_Down
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
