using Microsoft.Owin;
using NSwag.AspNet.Owin;
using Owin;
using System.Web.Http;
using DotNetEasyPdfSample.Web.DotNetFw;

[assembly: OwinStartup(typeof(Startup))]

namespace DotNetEasyPdfSample.Web.DotNetFw
{
    public class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            var configuration = new HttpConfiguration();
            configuration.MapHttpAttributeRoutes();

            app.UseSwaggerUi3(typeof(Startup).Assembly, settings =>
            {
                settings.Path = "/swagger";
            });

            app.UseSwaggerReDoc(typeof(Startup).Assembly, settings =>
            {
                settings.Path = "/redoc";
            });

            app.UseWebApi(configuration);
        }
    }
}
