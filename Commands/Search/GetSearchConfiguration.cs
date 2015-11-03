using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Administration;
using Microsoft.SharePoint.Client.Search.Portability;
using OfficeDevPnP.PowerShell.CmdletHelpAttributes;
using OfficeDevPnP.PowerShell.Commands.Enums;
using Encoding = System.Text.Encoding;
using File = System.IO.File;
using Resources = OfficeDevPnP.PowerShell.Commands.Properties.Resources;

namespace OfficeDevPnP.PowerShell.Commands.Search
{
    [Cmdlet(VerbsCommon.Get, "SPOSearchConfiguration")]
    [CmdletHelp("Returns the search configuration",
        Category = CmdletHelpCategory.Search)]
    [CmdletExample(
        Code = @"PS:> Get-SPOSearchConfiguration",
        Remarks = "Returns the search configuration for the current web",
        SortOrder = 1)]
    [CmdletExample(
        Code = @"PS:> Get-SPOSearchConfiguration -Scope Site",
        Remarks = "Returns the search configuration for the current site collection",
        SortOrder = 2)]
    [CmdletExample(
        Code = @"PS:> Get-SPOSearchConfiguration -Scope Subscription",
        Remarks = "Returns the search configuration for the current tenant",
        SortOrder = 3)]
    public class GetSearchConfiguration : SPOWebCmdlet
    {
        [Parameter(Mandatory = false)]
        public SearchConfigurationScope Scope = SearchConfigurationScope.Web;

        [Parameter(Mandatory = false, HelpMessage = "Filename to write to, optionally including full path")]
        public string Filename;

        [Parameter(Mandatory = false, HelpMessage = "Overwrites the output file if it exists.")]
        public SwitchParameter Force;

        [Parameter(Mandatory = false)]
        public Encoding Encoding = Encoding.Unicode;

        protected override void ExecuteCmdlet()
        {

            var xml = String.Empty;
            switch (Scope)
            {
                case SearchConfigurationScope.Web:
                    {
                        xml = SelectedWeb.GetSearchConfiguration();
                        break;
                    }
                case SearchConfigurationScope.Site:
                    {
                        xml = ClientContext.Site.GetSearchConfiguration();
                        break;
                    }
                case SearchConfigurationScope.Subscription:
                    {
                        //xml = ClientContext.GetSearchConfiguration(SearchObjectLevel.SPSiteSubscription);
                        break;
                    }
                //#if CLIENTSDKV15
                case SearchConfigurationScope.Ssa:
                    {
                        //xml = ClientContext.GetSearchConfiguration(SearchObjectLevel.Ssa);
                        break;
                    }
                    //#endif
            }

            if (!string.IsNullOrEmpty(Filename))
            {
                if (!Path.IsPathRooted(Filename))
                {
                    Filename = Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, Filename);
                }
                if (File.Exists(Filename))
                {
                    if (Force || ShouldContinue(string.Format(Resources.File0ExistsOverwrite, Filename), Resources.Confirm))
                    {
                        File.WriteAllText(Filename, xml, Encoding);
                    }
                }
                else
                {
                    File.WriteAllText(Filename, xml, Encoding);
                }
            }
            else
            {
                WriteObject(xml);
            }

        }
    }
}