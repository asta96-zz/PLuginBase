using Microsoft.Xrm.Sdk;
using System;

namespace Retro.Plugins
{
    public class UpdateonAccountName : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            tracing.Trace("context Depth:" + context.Depth);
            Entity target = (Entity)context.InputParameters["Target"];
            target["name"] = $"updating name via UpdateonAccountName.{Convert.ToString(context.Depth)}";
            if (context.Depth > 1)
            {
                return;
            }
            service.Update(target);
        }
    }
}