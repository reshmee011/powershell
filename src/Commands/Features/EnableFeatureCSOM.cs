using Microsoft.SharePoint.Client;
using PnP.PowerShell.Commands.Enums;
using System;
using System.Management.Automation;

namespace PnP.PowerShell.Commands.FeaturesCSOM
{
    [Cmdlet(VerbsLifecycle.Enable, "PnPFeature")]
    [OutputType(typeof(void))]
    public class EnableFeatureCSOM : PnPSharePointCmdlet
    {
        public EnableFeatureCSOM(Guid Identity, FeatureScope Scope)
        {
            // Default constructor
            // Ensure web/site properties loaded
            if (Scope == FeatureScope.Web)
            {
                ClientContext.Load(ClientContext.Web, w => w.Url);
            }
            else
            {
                ClientContext.Load(ClientContext.Site, s => s.Id);
            }
            ClientContext.ExecuteQueryRetry();

            // Select the appropriate FeatureCollection
            FeatureCollection feats = Scope == FeatureScope.Web ? ClientContext.Web.Features : ClientContext.Site.Features;

            // Load existing features
            ClientContext.Load(feats);
            ClientContext.ExecuteQueryRetry();

            try
            {
                // Prefer direct call if available
                feats.Add(Identity, false, 0);
                ClientContext.ExecuteQueryRetry();
            }
            catch (ServerException ex)
            {
                WriteError(new ErrorRecord(ex, "EnableFeatureFailed", ErrorCategory.InvalidOperation, null));
            }
        }
        protected override void ExecuteCmdlet()
        {
          
        }

    }
}
