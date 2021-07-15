<!-- markdownlint-disable MD002 MD041 -->

Microsoft Teams tab applications have multiple options to authenticate the user and call Microsoft Graph. In this exercise, you'll implement a tab that does [single sign-on](/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso) to get an auth token on the client, then uses the [on-behalf-of flow](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) on the server to exchange that token to get access to Microsoft Graph.

For other alternatives, see the following.

- [Build a Microsoft Teams tab with the Microsoft Graph Toolkit](/graph/toolkit/get-started/build-a-microsoft-teams-tab). This sample is completely client-side, and uses the Microsoft Graph Toolkit to handle authentication and making calls to Microsoft Graph.
- [Microsoft Teams Authentication Sample](https://github.com/OfficeDev/microsoft-teams-sample-auth-node). This sample contains multiple examples covering different authentication scenarios.

## Create the project

Start by creating an ASP.NET Core web app.

1. Open your command-line interface (CLI) in a directory where you want to create the project. Run the following command.

    ```Shell
    dotnet new webapp -o GraphTutorial
    ```

1. Once the project is created, verify that it works by changing the current directory to the **GraphTutorial** directory and running the following command in your CLI.

    ```Shell
    dotnet run
    ```

1. Open your browser and browse to `https://localhost:5001`. If everything is working, you should see a default ASP.NET Core page.

> [!IMPORTANT]
> If you receive a warning that the certificate for **localhost** is un-trusted you can use the .NET Core CLI to install and trust the development certificate. See [Enforce HTTPS in ASP.NET Core](/aspnet/core/security/enforcing-ssl?view=aspnetcore-3.1) for instructions for specific operating systems.

## Add NuGet packages

Before moving on, install some additional NuGet packages that you will use later.

- [Microsoft.Identity.Web](https://www.nuget.org/packages/Microsoft.Identity.Web/) for authenticating and requesting access tokens.
- [Microsoft.Identity.Web.MicrosoftGraph](https://www.nuget.org/packages/Microsoft.Identity.Web.MicrosoftGraph/) for adding Microsoft Graph support configured with Microsoft.Identity.Web.
- [Microsoft.Graph](https://www.nuget.org/packages/Microsoft.Graph/) to update the version of this package installed by Microsoft.Identity.Web.MicrosoftGraph.
- [TimeZoneConverter](https://github.com/mj1856/TimeZoneConverter) for translating Windows time zone identifiers to IANA identifiers.

1. Run the following commands in your CLI to install the dependencies.

    ```Shell
    dotnet add package Microsoft.Identity.Web --version 1.14.1
    dotnet add package Microsoft.Identity.Web.MicrosoftGraph --version 1.14.1
    dotnet add package Microsoft.Graph --version 4.0.0
    dotnet add package TimeZoneConverter
    ```

## Design the app

In this section you will create the basic UI structure of the application.

> [!TIP]
> You can use any text editor to edit the source files for this tutorial. However, [Visual Studio Code](https://code.visualstudio.com/) provides additional features, such as debugging and Intellisense.

1. Open **./Pages/Shared/_Layout.cshtml** and replace its entire contents with the following code to update the global layout of the app.

    :::code language="cshtml" source="../demo/GraphTutorial/Pages/Shared/_Layout.cshtml" id="LayoutSnippet":::

    This replaces Bootstrap with [Fluent UI](https://developer.microsoft.com/fluentui), adds the [Microsoft Teams SDK](/javascript/api/overview/msteams-client), and simplifies the layout.

1. Open **./wwwroot/js/site.js** and add the following code.

    :::code language="javascript" source="../demo/GraphTutorial/wwwroot/js/site.js" id="ThemeSupportSnippet":::

    This adds a simple theme change handler to change the default text color for dark and high contrast themes.

1. Open **./wwwroot/css/site.css** and replace its contents with the following.

    :::code language="css" source="../demo/GraphTutorial/wwwroot/css/site.css" id="CssSnippet":::

1. Open **./Pages/Index.cshtml** and replace its contents with the following code.

    ```cshtml
    @page
    @model IndexModel
    @{
      ViewData["Title"] = "Home page";
    }

    <div id="tab-container">
      <h1 class="ms-fontSize-24 ms-fontWeight-semibold">Loading...</h1>
    </div>

    @section Scripts
    {
      <script>
      </script>
    }
    ```

1. Open **./Startup.cs** and remove the `app.UseHttpsRedirection();` line in the `Configure` method. This is necessary for ngrok tunneling to work.

## Run ngrok

Microsoft Teams does not support local hosting for apps. The server hosting your app must be available from the cloud using HTTPS endpoints. For debugging locally, you can use ngrok to create a public URL for your locally-hosted project.

1. Open your CLI and run the following command to start ngrok.

    ```Shell
    ngrok http 5000
    ```

1. Once ngrok starts, copy the HTTPS Forwarding URL. It should look like `https://50153897dd4d.ngrok.io`. You'll need this value in later steps.

> [!IMPORTANT]
> If you are using the free version of ngrok, the forwarding URL changes every time you restart ngrok. It's recommended that you leave ngrok running until you complete this tutorial to keep the same URL. If you have to restart ngrok, you'll need to update your URL everywhere that it is used and reinstall the app in Microsoft Teams.
