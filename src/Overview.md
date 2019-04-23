
# SharePoint Build and Release Tasks

This extension includes a group of tasks that leverages SharePoint and O365 functionalities for build and release.

## Content:

1. #### [Task: Deploy PnP SharePoint Artifacts](#Task-Deploy-PnP-SharePoint-Artifacts)
2. #### [Change Log](#Change-Log)

## <a id="Task-Deploy-PnP-SharePoint-Artifacts"> </a> Task Deploy PnP SharePoint Artifacts

Deploys SharePoint artifacts (e.g. lists, fields, content type...) with the publish PnP PowerShell, which uses the PnP Provisioning Engine.
This task works mainly in the same way as described in the documentation of the [PnP PowerShell cmdlet Apply-PnPProvisioningTemplate](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/apply-pnpprovisioningtemplate?view=sharepoint-ps).
This PowerShell task allows you to use [PnP PowerShell](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp), which will be loaded prior executing any script. The newest releast modules are downloaded from the official PSGallery feed, if not present on the agent.

### Mandatory Fields

First the SharePoint version has to be chosen.

![SharePoint Choice](images/PnPDeploySpArtifacts/deploySpArtifacts01.png)

You need to create a service connection to SharePoint. This service connection comes with the extension. 

![Service Connection to SharePoint](images/PnPDeploySpArtifacts/deploySpArtifacts05.png)

> This connection is currently not working with MFA (multi factor authentication) enabled tenants. You need to either be inside the coporate network (or VPN connection) that allows authentication without MFA or you use a service account, that does not have MFA enabled.

After that use the created connection in your task

![Service Connection to SharePoint](images/PnPDeploySpArtifacts/deploySpArtifacts02.png)

Next, you choose if you want to use a file from your build or if you want to use inline xml. A [specific xml schema is expected](https://github.com/SharePoint/PnP-Provisioning-Schema/blob/master/ProvisioningSchema-2016-05.md) which is parsed by the PnP provisioning engine.

![Type of Input](images/PnPDeploySpArtifacts/deploySpArtifacts04.png)

You can include the following XML, which changes the title of the connected web

```xml
<?xml version="1.0"?>
<pnp:Provisioning xmlns:pnp="http://schemas.dev.office.com/PnP/2017/05/ProvisioningSchema">
  <pnp:Preferences Generator="OfficeDevPnP.Core, Version=2.18.1709.0, Culture=neutral, PublicKeyToken=3751622786b357c2" />
  <pnp:Templates ID="CONTAINER-TEMPLATE-CE97DA40966E445087F3E67032B06CC6">
    <pnp:ProvisioningTemplate ID="TEMPLATE-CE97DA40966E445087F3E67032B06CC6" Version="1" BaseSiteTemplate="STS#0" Scope="Web">
      <pnp:WebSettings NoCrawl="false" Title="My Web Title" WelcomePage="" AlternateCSS="" MasterPageUrl="{masterpagecatalog}/seattle.master" CustomMasterPageUrl="{masterpagecatalog}/seattle.master" />
    </pnp:ProvisioningTemplate>
  </pnp:Templates>
</pnp:Provisioning>
```

### Optional Fields

#### Handler To Be Used

Then you can optionally give a comma separated list of Handlers (e.g. Lists,Fields). Leave empty if all Handlers should be used. This Allows you to only process a specific part of the template. Notice that this might fail, as some of the handlers require other artifacts in place if they are not part of what your applying. Check for [available Handlers.](https://msdn.microsoft.com/en-us/pnp_sites_core/officedevpnp.core.framework.provisioning.model.handlers)

#### Parameters To Be Added

The field "Parameters To Be Added" allows you to specify parameters that can be referred to in the template by means of the {parameter:} token. use only one parameter-value pair per line.

Example:

```dictionary
ListTitle=Projects
parameter2=a second value
```

![Parameters](images/PnPDeploySpArtifacts/deploySpArtifacts03.png)

See examples on [how it works internally](https://github.com/SharePoint/PnP-PowerShell/blob/master/Documentation/ApplyPnPProvisioningTemplate.md#example-3).

### Advanced Parameters

#### ClearNavigation
Override the RemoveExistingNodes attribute in the Navigation elements of the template. If you specify this value the navigation nodes will always be removed before adding the nodes in the template.

#### Ignore Duplicate Data Row Errors
Ignore duplicate data row errors when the data row in the template already exists.

#### Overwrite System Property Bag Values
Specify this parameter if you want to overwrite and/or create properties that are known to be system entries (starting with vti_, dlc_, etc.)

#### Provision Content Types To Sub Webs
If set content types will be provisioned if the target web is a subweb.



---
# <a id="Change-Log"> </a> Change Log

## 1.0.0

### Mayor changes

- Some mayor changes

### Minor

- Some minor changes

### Bug Fixes

- Some bug fixes 


