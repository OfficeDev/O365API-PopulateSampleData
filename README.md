---
topic: sample
products:
- office-onedrive
- office-365
languages:
- csharp
extensions:
  contentType: samples
  createdDate: 2/26/2015 2:49:40 PM
---
# Populate Office 365 Developer Tenant With Sample Data
A Windows Store App that will populate data for Office 365 API services such as mail, calendar, contact and files. The app essentially reads new contacts, mails and calendar events from an XML file and add them to the logged in Office 365 user's account. Existing data on the Office 365 tenant account will not be affected by this app.

You can use this app to quickly populate data into your Office 365 developer tenant so you can get started buidling Office 365 apps by interacting with the populated data. 

You can find the XML files for contacts, events and mail under the Assets folder:
- [AddContact.xml](https://github.com/OfficeDev/O365API-PopulateSampleData/blob/master/O365DataApp/O365DataApp/Assets/AddContact.xml)
- [AddEvent.xml](https://github.com/OfficeDev/O365API-PopulateSampleData/blob/master/O365DataApp/O365DataApp/Assets/AddEvent.xml)
- [AddMail.xml](https://github.com/OfficeDev/O365API-PopulateSampleData/blob/master/O365DataApp/O365DataApp/Assets/AddMail.xml)

For Files, upload your documents to a folder of your choice in your development machine and use the app to select one or more of the documents to upload.

## How To Run This Sample
To run this sample, you need:

1. Visual Studio 2013
2. [Office Developer Tools for Visual Studio 2013](http://aka.ms/OfficeDevToolsForVS2013)
3. Office 365 Developer Subscription. [Join the Office 365 Developer Program and get a free 1 year subscription to Office 365](https://aka.ms/devprogramsignup).

## Step 1: Clone the application in Visual Studio
Visual Studio 2013 supports connecting to Git servers. As the project templates are hosted in GitHub, Visual Studio 2013 makes it easier to clone projects from GitHub.

The steps below will describe how to clone Office 365 API web application project in Visual Studio from Office Developer GitHub.

1. Open Visual Studio 2013.
2. Switch to Team Explorer.
3. Team Explorer provides options to clone Git repositories.
4. Click Clone under Local Git Repositories, enter the clone URL **https://github.com/OfficeDev/O365API-PopulateSampleData.git** for the web application project and click Clone.
5. Once the project is cloned, double click on the repo.
6. Double click the project solution which is available under Solutions.
7. Switch to Solution Explorer.

## Step 2: Configure the sample

### Build the Project
Simply Build the project to restore NuGet packages.

### Register Azure AD application to consume Office 365 APIs
Office 365 applications use Azure Active Directory (Azure AD) to authenticate and authorize users and applications respectively. All users, application registrations, permissions are stored in Azure AD.

Using the Office 365 API Tool for Visual Studio you can configure your application to consume Office 365 APIs.

1. In the Solution Explorer window, **right click your project -> Add -> Connected Service**.
2. A Services Manager dialog box will appear. Choose **Office 365 -> Office 365 API** and click **Register your app**.
3. On the sign-in dialog box, enter the username and password for your Office 365 tenant.
4. After you're signed in, you will see a list of all the services.
5. Initially, no permissions will be selected, as the app is not registered to consume any services yet.
6. Select **Users and Groups** and then click **Permissions**
7. In the **Users and Groups Permissions** dialog, select **Enable sign-on and read users profiles'** and click **Apply**
8. Select **My Files** and then click **Permissions**
9. In the **My Files Permissions** dialog, select both **Read users' files** and **Edit or delete users' files** then click **Apply**
10. Select **Mail** and then click **Permissions**
11. In the **Mail Permissions** dialog, select both **Read and write access to users' mail** and **Send mail as a user** then click **Apply**
12. Select **Contacts** and then click **Permissions**
13. In the **Contacts Permissions** dialog, select **Have full access to users' contacts** and click **Apply**
14. Select **Calendar** and then click **Permissions**
13. In the **Calendar Permissions** dialog, select **Have full access to users' calendars** and click **Apply**
11. Click **Ok**

After clicking OK, Office 365 client libraries (in the form of NuGet packages) for connecting to Office 365 APIs will be added to your project.

In this process, Office 365 API tool registered an Azure AD Application in the Office 365 tenant that you signed in the wizard and added the Azure AD application details to App.xaml.

### Set Contacts
The application reads contacts' details from a .xml file and adds each contact to the Office 365 tenant account.

1. Open Assets\AddContact.xml file 
2. For your reference, 2 sample contacts already exist on the file.
3. Add as many **Contact** xml tags along with rest of the xml tags that correspond to each contact on this file.
4. Save changes

**Note:** Each contact added to this file will be added to the Office 365 tenant account.

### Set Calendar Events
The application reads calendar events' details from a .xml file and adds each event to the Office 365 tenant account.

1. Open Assets\AddEvent.xml file 
2. For your reference, 2 sample events already exist on the file.
3. Add as many **Event** xml tags along with rest of the xml tags that correspond to each event on this file.
4. Save changes

**Note:** Each event added to this file will be added to the Office 365 tenant account.

### Set Mails
The application reads mails' details from a .xml file and adds each mail to the Office 365 tenant account.

1. Open Assets\AddMail.xml file 
2. For your reference, 2 sample messages already exist on the file.
3. Add as many **Message** xml tags along with rest of the xml tags that correspond to each message on this file. **Note:** Make sure that your Office 365 tenant email address is on the ToRecipient list.
4. Save changes

**Note:** Each message added to this file will be added to the Office 365 tenant account.

## Step 3: Run the sample App

Press F5, to run the app.

This is how the application looks:

![Office 365 Connect sample](/readme-images/O365DataApp.jpg "O365DataApp Home Screen")

### Add Files to OneDrive
1. Click on the **Add MyFiles** button
2. If prompted for user credentials, then enter in the username and password for the same Office 365 tenant account to which the app was registered, click **Sign in** and accept the permissions request.
3. Navigate to the files' location.
4. Select all files that needs to be added.
5. Click Open.

Once the files have been added to your office 365 tenant account, the names of the files will be displayed on the home screen on the app.

### Add Contacts
1. Click on **Contacts** button on the app.
2. If prompted for user credentials, then enter in the username and password for the same Office 365 tenant account to which the app was registered, click **Sign in** and accept the permissions request.

The contacts added on **AddContact.xml** file will be added to your office 365 tenant account, the display name for each contact will be displayed on the home screen on the app.

### Add Calendar Events

1. Click on **Events** button on the app.
2. If prompted for user credentials, then enter in the username and password for the same Office 365 tenant account to which the app was registered, click **Sign in** and accept the permissions request.

The events added on **AddEvent.xml** file will be added to your office 365 tenant account, the subject for each event will be displayed on the home screen on the app.

### Add Mails

1. Click on **Mails** button on the app.
2. If prompted for user credentials, then enter in the username and password for the same Office 365 tenant account to which the app was registered, click **Sign in** and accept the permissions request.

The messages added on **AddMail.xml** file will be sent to the specified To/Cc/Bcc recipients.

### Update Contacts
1. If you are already running the app then stop running the app.
2. Open **Assets\AddContact.xml** file in VS from solution explorer.
3. Change **ActionName** to **UPDATE**.
4. Make the necessary changes to each contact and save all changes on the file. **Note:** To update an already existing contact, the application will look for the contact in your office 365 account with the same **GivenName** and **Surname**. Thus you can update all other properties of a contact apart from their GivenName and Surname.
5. Press F5
6. Click on **Contacts** button on the app.

### Update Calendar Events
1. If you are already running the app then stop running the app.
2. Open **Assets\AddEvent.xml** file in VS from solution explorer.
3. Change **ActionName** to **UPDATE**.
4. Make the necessary changes to each event and save all changes on the file. **Note:** To update an already existing event, the application will look for the event in your office 365 account with the same **Subject**. Thus you can update all other properties of an event apart from their Subject.
5. Press F5
6. Click on **Events** button on the app.

### Delete Contacts
1. If you are already running the app then stop running the app.
2. Open **Assets\AddContact.xml** file in VS from solution explorer.
3. Change **ActionName** to **DELETE**.
4. Press F5
5. Click on **Contacts** button on the app.

**Note:** The app will look for contacts with the specified GivenName and Surname and delete them.

### Delete Calendar Events
1. If you are already running the app then stop running the app.
2. Open **Assets\AddEvent.xml** file in VS from solution explorer.
3. Change **ActionName** to **DELETE**.
4. Press F5
5. Click on **Events** button on the app.

**Note:** The app will look for events with the specified Subject and delete them.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
