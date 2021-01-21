# OutlookSpamButton
Module in VBA to add as a "Report Spam" button in outlook and send it to verdacht@safeonweb.be

## How it works
This module will send the current selected email in outlook as an attachment to verdacht@safeonweb.be and will delete the email from your inbox. This can be added as a button in your outlook ribbon for easy use.

## How to install

### Step 1: Enable developer mode
Open Outlook and navigate to **File > Options**.  
Under **Customize Ribbon**, in the **Main Tabs** list, check the **Developer** box if it is not already checked.
![Outlook Options](https://github.com/T0nyMacaroni/OutlookSpamButton/blob/main/images/1-outlookoptions.png?raw=true) 

### Step 2: Import module
In the outlook ribbon select the new **Developer** tab and click on **Visual Basic** (or press Alt + F11)  
In the Visual Basic window click on File > import module and import the module from this repository.  
![Import Module](https://github.com/T0nyMacaroni/OutlookSpamButton/blob/main/images/2-importmodule.png?raw=true)  

### Step 2b: Change email address (optional)
Change the emailaddress in the code that you would like to send your spam to.  
By default, this mail will be sent to verdacht@safeonweb.be. Feel free to change this.

### Step 3: Add button in ribbon
Pick the tab where you want to add your own group. 

Select **New Group**.  
That adds **New Group (Custom)** to the tab you picked.

To use a better name for your new group, click **Rename**, type the name you want in the **Display name** box, and then click **OK**.  

To add a macro to the group, in the **Choose commands from** list, click **Macros**.

Select the macro you imported to add to your new group, and then click **Add**. The macro is added to your **Custom group**.

![Import Module](https://github.com/T0nyMacaroni/OutlookSpamButton/blob/main/images/3-customizeribbon.png?raw=true) 

To use a friendlier name, click **Rename**, and then type the name you want in the **Display name** box and choose an icon. (optional)

![Rename](https://github.com/T0nyMacaroni/OutlookSpamButton/blob/main/images/4-rename.png?raw=true) 

Click **OK** to save.

The new group appears on the tab you picked, where you can click the button to run the macro. 