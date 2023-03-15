[![Hack Together: Microsoft Graph and .NET](https://img.shields.io/badge/Microsoft%20-Hack--Together-orange?style=for-the-badge&logo=microsoft)](https://github.com/microsoft/hack-together)

 # .NET MVC (Model - View - Controller) web app connected to Microsoft Graph

This is a template for .NET Core MVC (Model - View - Controller) web app that connects to Microsoft Graph.

# Introduction

**Application name** : Microsoft 365 Tiles

The application provides different Microsoft 365 tiles and charts based on the Microsofrt Graph API for the currently logged in user. Application has been built using .Net 7.0 MVC and with Microsoft Graph API SDK Version **4.2.0**

The application has the following main features

* > **Charts**
    1. Messages stats
    2. Chat stas
    3. Files stats
    4. Tasks stats
* > **Tiles**
    1. Recent documents
    2. Viewed and modified
    3. OneDrive items
    4. Notebooks
    5. Upcoming events
    6. Followed sites
    7. Contacts
    8. Colleagues around
    9. Recent tasks
    10. Joined Teams
    11. Teams stats

# Demo

![demo](./Assets/ms-graph-tiles-demo.gif)

# Screenshots

### Charts
![screenshot-charts](./Assets/screenshot-charts.png)

### Tiles
![screenshot-tiles](./Assets/screenshot-tiles.png)

## Minimal Path to Awesome 🚀

Follow the instructions to successfully run your MVC app with Microsoft Graph. 
### 1. Register an Azure Active Directory app
Please follow the below steps to register your app in Azure AD: 

* Go to [Azure Portal](https://portal.azure.com) and login with your testing account that has Application developer or administrator permissions.
* Select **Azure Active Directory**, and select **App Registrations** from the left side bar. Then select **+ New registration**.
* Give any name to your app. For **Supported account types**, select **Azure AD and Personal Accounts (Any Azure AD directory - Multitenant)**.
* Select **Authentication** from the left navigation, and select **Add a platform**
* Select **Web** and enter https://localhost:7004/signin-oidc as Redirect URI (replace the port number with your local host port)
* Make sure **Access tokens (used for implicit flows)** and **ID tokens (used for implicit and hybrid flows)** are selected
* Select **Save**
* Navigate to **Certificates & secrets** and create a new client secret, make a note of the client secret
* Select **API permissions* and select **Add a permission**, then add the following delegated permissions:

![permissions](./assets/permissions.png)

* Set the **Redirect URI** drop down to **Public client/native (mobile & desktop)** and enter `http://localhost`. Then, select **Register**. Navigate to **Overview** tab and make a note of the **Application (client) ID**. You'll use it in the next steps. </details> ### 2. Run your Console app * Clone the Hack Together repository to your local workspace or directly download the source code.
* Open the project folder `dotnet-console-app-microsoft-graph` with the editor of your choice. (Visual Studio Code is recommended.)
* In Visual Studio Code, press F5 to run the app.
  ![App output in the debug console in VSCode](./screenshot.png) ## Reference * [Quickstart: Register an application with the Microsoft identity platform](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
