# Microsoft Teams Presence Information Using Microsoft Graph API
This is a sample react app to showcase Teams presence information using Graph API
## Prerequisites
- A text editor or IDE. You can install and use Visual Studio Code for free.


## Register an app in Azure portal

In this step you'll register an application in the Azure AD admin center. This is necessary to authenticate the application to make calls to the Microsoft Graph indexing API.

1. Go to the [Azure Active Directory admin center](https://aad.portal.azure.com/) and sign in with an administrator account.
1. Select **Azure Active Directory** in the left-hand pane, then select **App registrations** under **Manage**.
1. Select **New registration**.
1. Complete the **Register an application** form with the following values, then select **Register**.

    - **Name:** `Flight Monitoring System`
    - **Supported account types:** `Accounts in this organizational directory only (Microsoft only - Single tenant)`
    - **Redirect URI:** Leave blank

1. On the **Flight Monitoring System** page, copy the value of **Application (client) ID**, you'll need it in the next section.
1. Copy the value of **Directory (tenant) ID**, you'll need it in the next section.
1. Select **API Permissions** under **Manage**.
1. Select **Add a permission**, then select **Microsoft Graph**.
1. Select **Application permissions**, then select the **Sites.Read.All, Sites.ReadWrite.All** permission. Select **Add permissions**.
1. Select **Grant admin consent for {TENANT}**, then select **Yes** when prompted.
