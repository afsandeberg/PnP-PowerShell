using System;
using System.Management.Automation;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Administration;
using OfficeDevPnP.PowerShell.CmdletHelpAttributes;
using OfficeDevPnP.PowerShell.Commands.Enums;
using Resources = OfficeDevPnP.PowerShell.Commands.Properties.Resources;

namespace OfficeDevPnP.PowerShell.Commands.Search
{
    [Cmdlet(VerbsCommon.Set, "SPOSearchConfiguration", DefaultParameterSetName = "SettingsString")]
    [CmdletHelp("Sets the search configuration",
        Category = CmdletHelpCategory.Search)]
    [CmdletExample(
        Code = @"PS:> Set-SPOSearchConfiguration -Configuration $config",
        Remarks = "Sets the search configuration for the current web",
        SortOrder = 1)]
    [CmdletExample(
        Code = @"PS:> Set-SPOSearchConfiguration -Configuration $config -Scope Site",
        Remarks = "Sets the search configuration for the current site collection",
        SortOrder = 2)]
    [CmdletExample(
        Code = @"PS:> Set-SPOSearchConfiguration -Configuration $config -Scope Subscription",
        Remarks = "Sets the search configuration for the current tenant",
        SortOrder = 3)]
    public class SetSearchConfiguration : SPOWebCmdlet
    {
        [Parameter(Mandatory = true, ParameterSetName = "SettingsString", Position = 0, ValueFromPipelineByPropertyName = true, ValueFromPipeline = true)]
        public string Configuration;

        [Parameter(Mandatory = true, ParameterSetName = "SettingsFile", Position = 0, ValueFromPipelineByPropertyName = true, ValueFromPipeline = true, HelpMessage = "Path to the xml file containing the search settinigs.")]
        public string Path;

        [Parameter(Mandatory = false)]
        public SearchConfigurationScope Scope = SearchConfigurationScope.Web;

        protected override void ExecuteCmdlet()
        {
            if (ParameterSetName.Equals("SettingsFile", StringComparison.InvariantCultureIgnoreCase))
            {
                Path = !System.IO.Path.IsPathRooted(Path) ? System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, Path) : Path;
                Configuration = System.IO.File.ReadAllText(Path);

            }

            switch (Scope)
            {
                case SearchConfigurationScope.Web:
                    {
                        this.SelectedWeb.SetSearchConfiguration(Configuration);
                        break;
                    }
                case SearchConfigurationScope.Site:
                    {
                        ClientContext.Site.SetSearchConfiguration(Configuration);
                        break;
                    }
                case SearchConfigurationScope.Subscription:
                    {
                        ClientContext.ImportSearchSettings(Configuration, SearchObjectLevel.SPSiteSubscription);
                        break;
                    }
                case SearchConfigurationScope.Ssa:
                    {
                        ClientContext.ImportSearchSettings(Configuration, SearchObjectLevel.Ssa);
                        break;
                    }
            }
        }
    }
}
