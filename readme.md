# MS Graph Workshop

Great that you chose to dig deeper into Microsoft Graph. This Lab instructions provide you with all information on how you can enable a Microsoft Graph based application.

This lab is completely implemented in TypeScript, however all the concepts presented are based on REST and therefore completely programming language agnostic. To use this lab, you have to execute three steps: Familiarize yourself with the App and with the bahavior of MS Graph, register your own Application in Azure Active Directory and implement the right queries to bring the App up and running.

## Set up the lab

To get started, fork or clone this repository to your computer. You will need Node.js installed to run this project, all other developer dependencies are delivered through the package.json. All dependencies are used for general development convenience, the MS Graph implementation is only dependent on the request module (but could potentially be implemented with the native Node https module as well).

## Get to know the App

If you are not taking this lab on your own, your instructor  most likely provided a running application as well as an user account to access that application to you.

If you do not have such an application already up and running, you can set up an application yourself. To do so, you will need to do three things: First, switch to the Solution branch in this repository and clone the repository. Second, set up an Azure Active Directory as described in the section *Setting up an Azure Active Directory Application* and set up a .env file like described. Then type in `npm start` in a console at the root of the repository.

When you first open the app at `http://localhost:3500` (or any other URL provided by the instructor), you will be redirect to the Microsoft login page. Sign in with the user information provided by your instructor. You will then be forwarded to a dashboard. Here you can see your email adress, your name, your user id and a field for your phone number. 

You can click on "Enable phone number editing" and on "Enable additional permissions". If you're clicking on these buttons, they will (with exception of the Tenant KPIs) redirect you to the Microsoft login page where you will be asked for your consent on that the application can read and write additional data on your behalf. This behavior is called **delegated permissions** where the application only acts in the name of a certain user and has at most the permissions that one user would have.

This is called **dynamic consent** in Azure Active Directory, where an application gets only minimal permissions for the beginning and requests additional permissions at the time needed.

When you click on "Enable additional permissions" beneath the Tenant KPIs, the App will log in and crawl Tenant data. It is now utilizing **Application permissions**, meaning that it acts on its own and gets tenant-wide access to data. This only works if your instructor (in his role as an administrator) has granted the application certain priviledges beforehand.

## Setting up an Azure Active Directory Application

To access the Microsoft Graph, you will always need an Azure Active Directory Application that handles permission management for you. To register a new application, visit portal.azure.com and log in with an user that is part of the tenant in which you want to use the application. After successful login, navigate to *Azure Active Directory* (usually in the favorites section to the left, otherwise you can search for it). Here click on *App registrations* and then *New application registration*. Give your app a name and a Sign-On URL (enter http://localhost:3500/api/callback here). Then click on create. In the upcoming window, save the *Application ID* (also called Client ID) in a text document. We will need that later. Then click on *Settings* and from here on *Keys*. Create a new password and save it to the text document (Make sure to do this as you won't be able to view the password again in the portal!).

Your app is (for now) completely set up. Go back to the source code on your computer and add a .env file in the root directory of the code. It should have the following structure:

````
TENANT_ID=<yourtenant>.onmicrosoft.com
CLIENT_ID=
CLIENT_SECRET=
REDIRECT_URL=http://localhost:3500/api/callback
````

## Connect your app to MS Graph

When starting the app now, you will get to the Microsoft login page. But the app will not be able to load as it has to query the users basic profile data. So let's get into implementing our MS Graph connectors. Note: It is your responsibility to find the right endpoints to query the data. Use https://docs.microsoft.com/en-us/graph/api/overview
to look for the suiting endpoints.
1. To get the app to loading, you have to implement the get() function in *src -> Entities -> Graph -> User.ts*
2. To enable phone number editing, implement the update() function in *src -> Entities -> Graph -> User.ts*
3. To enable User KPIs, implement:
    * get() in *src -> Entities -> Graph -> CalendarEvent.ts*
    * getSentMailCount() in *src -> Entities -> Graph -> MailBox.ts*
4. To enable Tenant KPIs:
    1. Request application permissions for your app in the Azure Portal. Go back to your Application Registration in Azure Active Directory and click on *Required permissions*. Click *Microsoft Graph* and select the following permissions (Be very precise and take exactly the permissions specified, not the readWrite alternative):
        * Read files in all site collections **(files.read.all)**
        * Read all users' full profiles **(user.read.all)**
        * Read all groups **(group.read.all)**
        * Read mail in all mailboxes **(mail.read)**
        * Read all audit log data **(auditlog.read.all)**
    2. Inform your instructor that you have set Application permissions in the Azure Active Directory Portal (he will then grant you these permissions in the portal).    
    3. Implement these calls:
        * getAllUsers() in *src -> Entities -> Graph -> User.ts*
        * get() in *src -> Entities -> Graph -> AuditLogs.ts*

## Further resources
If you followed all instructions, your app should now be up and running. But most likely your instructor will shut down your account and the app after the workshop, so how to practice further without messing around with your company tenant? 

Microsoft provides the Office Developer Program (https://developer.microsoft.com/en-us/office/dev-program). Included in this program is an one-year Office subscription with a tenant and 25 E3 licences where you can set up users and experiment with nearly all features of Microsoft Graph in a safe environment.