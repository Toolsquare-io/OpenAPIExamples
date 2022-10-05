# OpenAPI Examples

We are moving to an Open API to make it possible to create a variety of integration like reservations, invoicing,… To start we’ll have to begin with user management. So in this release (oktober 2022) you’ll be able to create users, change their user group and delete them if needed. To get you started we’ve created an example application using google sheets and the google script platform, but feel free to create your own application.

## Google Script Setup

Create a new google sheet. In the menubar go to Extensions -> Apps Script. That will open up a new tab where you'll see a first function. Here we have to do a couple of things:

### Add the code

Delete the function in the code panel and replace it with [this code](https://raw.githubusercontent.com/Toolsquare-io/OpenAPIExamples/main/UserDemo.gs).

### Add the OAuth2 libary
In the google script page you'll see a library tab with a + sign. Click that and add [this library](https://github.com/googleworkspace/apps-script-oauth2) using this ScriptID:

```
1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
```
If paste that ID in the AddLibary box and click **Lookup** you should see version 42 or more and OAuth2 below, click **add**.

### Get your credentials

To allow the script to talk to your tenant we have to generate you a ClientID and a ClientSecret. That's something we can create for you. We will also need the script ID so the authentication knows where to redirect the authentication to. So in the google script tab you'll see a gear on the left. There you can copy the script ID. If you send that to us we'll send you back a linked ClientID and ClientSecret. Once generated we can’t see them anymore, so make sure you store them well.

Now you'll have to save your script (the disk icon on top) and you can close the tab.

## First Run

If you refresh your google sheet, you should see a Toolsquare menu (could take a coupel of seconds to load). If you now select the **Create User Template** you'll be asked for the tenant and those ClientID and ClientSecret. The tenant is this part in your url: *tenant*.toolsquare.io.

If you added those values you should see a menu poping up with a link. That should bring you to an OAuth authentification page. If you're already loggedin in the Toolsquare platform you should **agree** with sharing data. If you now rerun the **Create User Template** the script will create a new spreadsheet tab with headers like this:

| Fist Name  | Last Name  | Email  | Usergroup|
|---|---|---|---|
|   |   |   |   |
|   |   |   |   |

Below the usergroup you should see a dropdown where you can select one of the available usergroups on your platform. When you see these, you know that the sheet is linked to the platform! 

## Using the sheet
Now you can copy/paste a list of users in the sheet according to the layout. Make sure the usergroups already exist on the platform (so are available in the dropdown). Then select the **Create Users** from the Toolsquare menu. All new users will receive an invitation to activate their account on the platform. The script will also **add** (for new users) or **change** (for existing users) the user group.


