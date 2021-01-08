using Microsoft.Xrm.Sdk;
using System;

namespace Retro.Plugins
{
    public class UpdateAccountDescription : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            tracing.Trace("context Depth:" + context.Depth);
            ////Entity target = (Entity)context.InputParameters["Target"];
            ////target["description"] = $"updating description via UpdateAccountDescription with Depth:{Convert.ToString(context.Depth)} at:{ DateTime.Now}";
            ////if(context.Depth>1)
            ////{
            ////    return;
            ////}Retro@1111
            ////service.Update(target);
            ///

            if (context.PrimaryEntityName.Equals("account", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidPluginExecutionException(OperationStatus.Failed, "Custom exception on update of Description");
                Entity entity = (Entity)context.InputParameters["Target"];
                entity["description"] = "updated in Preoperation stage";
            }
        }
    }
}