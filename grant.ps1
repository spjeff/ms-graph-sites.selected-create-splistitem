# from https://ashiqf.com/2021/03/15/how-to-use-microsoft-graph-sharepoint-sites-selected-application-permission-in-a-azure-ad-application-for-more-granular-control/
Connect-PnPOnline "https://spjeff.sharepoint.com"
Grant-PnPAzureADAppSitePermission -AppId "fe796649-2134-496e-8fd6-bbd90ff12a01" -DisplayName "SPListWrite" -Permissions "Write"
