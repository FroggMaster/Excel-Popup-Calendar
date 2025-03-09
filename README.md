# Excel VBA Calendar Integration with Mini Calendar and Date Picker Add-in

## Overview
This project integrates an **Excel VBA macro** to toggle the visibility of a **Mini Calendar and Date Picker Add-in** in Excel. When a specified date cell is selected, the calendar pops up, allowing the user to pick a date. The calendar visibility is toggled based on the selection, and the previous date value is stored and restored if needed.

---

## Prerequisites
- **Microsoft Excel** with support for VBA macros.
- **Mini Calendar and Date Picker Add-in** installed in Excel.
  - Install it from [AppSource](https://appsource.microsoft.com/en-us/product/office/WA102957665?src=office&corrid=59aee7e2-a7ab-e564-e436-39ba83d0c3ba&omexanonuid=&referralurl=&ClientSessionId=19a83f6d-f232-4e4f-ac0a-1852c31fabe3) or from the Developer tab in Excel.
---

<details>
    <summary> Click to expand Install Instructions for Mini Calendar and Date Picker Add-in </summary>

### **Install the Mini Calendar and Date Picker Add-in:**

<details>
  <summary>Install the Mini Calendar and Date Picker Add-in VIA Microsoft AppSource</summary>
  
1. **Install the Add-in:**
   - Navigate to the [Mini Calendar and Date Picker Add-in](https://appsource.microsoft.com/en-us/product/office/WA102957665?src=office&corrid=59aee7e2-a7ab-e564-e436-39ba83d0c3ba&omexanonuid=&referralurl=&ClientSessionId=19a83f6d-f232-4e4f-ac0a-1852c31fabe3) page on Microsoft AppSource.
   - Click **Get It Now** to install the add-in into your Excel application.
   - After installation, ensure the add-in is available by navigating to **Insert > My Add-ins**.
</details>

<details>
    <summary>Install the Mini Calendar and Date Picker Add-in VIA Developer Mode</summary>

1. **Enable Developer Tab:**
   - Go to the **File** tab and click on **Options** to open the **Excel Options** window.
   - In the **Excel Options** window, select **Customize Ribbon** from the left sidebar.
   - In the **Customize the Ribbon** section, check the box labeled **Developer** under the **Main Tabs** list.
   - Click **OK** to confirm and close the options window.
   - Now the **Developer** tab should appear in the Excel ribbon.

2. **Open the Developer Tab and Access Add-ins:**
   - Click on the **Developer** tab in the ribbon at the top of Excel.
   - In the **Developer** tab, click on **Add-ins** to open the available controls.
   - In the **Add-ins** section, click on **Store**. This will open the **Office Add-ins** dialog box.

3. **Search for the Mini Calendar and Date Picker Add-in:**
   - In the **Office Add-ins** dialog box, type **Mini Calendar and Date Picker** into the search bar.
   - The add-in should appear in the search results.
   - Click on **Add** to install the add-in.

4. **Install the Add-in:**
   - Once you click **Add**, the **Mini Calendar and Date Picker Add-in** will be installed and added to Excel.
   - You should see a new button for the add-in in the **Add-ins** section.
</details>


</details>


---
1. **Create a Popup Calendar:**
   - Click on the **Developer** tab in the ribbon at the top of Excel.
   - In the **Developer** tab, click on **Add-ins** to open the available controls.
   - Click the button for the **Mini Calendar and Date Picker Add-in** to create the calendar.
   - Resize the calendar as necessary.
   - Press **Alt + F10**: This will open the **Selection Pane**, where you can see the names of all objects on the current worksheet.
   - Rename the calendar object to "Calendar" or any name you prefer, but you will need to adjust the VBA code to match the object’s name.

2. **Using the VBA Code:**
   - Press **Alt + F11** to open the **Visual Basic for Applications (VBA) editor** in Excel.
   - In the **Project Explorer** window (on the left side), locate the worksheet you want to apply the code to (e.g., `Sheet1`).
   - Double-click the worksheet name to open its code window.
   - Copy and paste the VBA code from [Popup_Calendar_Worrksheet_Code.cls](https://github.com/FroggMaster/Excel-Popup-Calendar/blob/main/Popup_Calendar_Worrksheet_Code.cls) into the worksheet’s code window.

3. **Modify the VBA Code**
   - You will need to set the calendar name to match the calendar object you created in variable `calendarName`. [Popup_Calendar_Worrksheet_Code.cls - Line 8](https://github.com/FroggMaster/Excel-Popup-Calendar/blob/main/Popup_Calendar_Worrksheet_Code.cls#L8)
   - You will need to set the cell or range where you want popup calendars to activate in variable `rngDate`. [Popup_Calendar_Worrksheet_Code.cls - Line 11](https://github.com/FroggMaster/Excel-Popup-Calendar/blob/main/Popup_Calendar_Worrksheet_Code.cls#L11)
