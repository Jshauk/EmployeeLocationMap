# Employee Location Map

This project is a SharePoint Framework (SPFx) web part paired with Power Automate flows that together create a dynamic, searchable employee directory and interactive seating map. Employee data is synchronized from Microsoft Entra ID (formerly Azure AD) into a SharePoint list, and seating locations are displayed using an SVG floor plan.

![Working Demo](/doc/demo.gif)


---

# Workflow

This application begins with a scheduled Power Automate flow that runs once per day.  
The workflow performs the following:

1. Reads employee information from Entra ID, including the "office" attribute and phone number.
2. Iterates through a security group containing all active employees.
3. Compares Entra data to the existing SharePoint employee list.
4. Removes terminated employees, adds new employees, and updates employee records as needed.
5. Works in both hybrid and cloud-only environments because data comes directly from Entra ID. This allows employee locations to be updated through on-prem AD (synced to Entra) or directly in Entra ID.

The imported flow file `UpdateEmployeeList.zip` contains this workflow.

To generate the initial data set, the `GenerateEmployeeList.zip` flow can be imported and run once to create the full SharePoint list of employees.

---

# Prerequisites

1. You must generate the initial SharePoint employee list. This can be done using the provided flow in `GenerateEmployeeList.zip`.

2. You must create an SVG map of your seating layout.
   - Use an SVG editing tool such as Inkscape.
   - Each cubicle, desk, or office must have its `id` attribute set to match the user's `office` attribute in Entra ID.
   - For example, a seat ID of `B10` in the SVG must correspond to a user's `office` value of `B10` in Entra or AD.

   ![Example of Svg map](https://github.com/Jshauk/EmployeeLocationMap/blob/main/doc/svg%20example.png)
   ![inkscape svg doc](https://github.com/Jshauk/EmployeeLocationMap/blob/main/doc/svg%20example%202.png)

3. This project is designed to store and load SVG maps from a SharePoint site. Storing maps on SharePoint allows updates to the floor plan without needing to repackage the SPFx solution.

4. The sample logic assumes two floors.
   - In the example environment, seat colors are used to differentiate floors.
   - Blue and Pink locations are found on floor 4; all other seat colors are on floor 3.
   - Therefore, a seat such as `B10` falls on floor 4 in the example configuration.
   Your environment may use different naming or categorization logic.

5. To move an employee to a new desk, update the user’s `office` attribute in Entra ID (or in AD if syncing). The daily flow will automatically update the SharePoint list and the map will reflect the change.

---

# Configuration for Use

Modify the following files to match your own environment.

### 1. config/serve.json
Update the `initialPage` value to point to your tenant’s SharePoint Workbench URL for testing.


### 2. src/webparts/listDirectory/ListDirectoryWebPart.manifest.json
Update the following properties:

- `iconImageUrl`
- `fullPageAppIconImageUrl`

These should be URLs to your app icons hosted in your SharePoint tenant or stored locally in your project.

### 3. src/webparts/listDirectory/components/ListDirectory.tsx
Update these variables:

- `targetSiteUrl`  
  The SharePoint site that contains the employee list.

- `svgFile`  
  The URL of the SVG seating map stored in your SharePoint site.

- `profilePictureUrl`  
  The base URL used to retrieve user profile photos in your tenant.

All of these must be updated to correspond to your tenant and list configuration.

---

# Notes and Limitations

- This solution has been tested with fewer than 200 employees.
- The SPFx component does not currently implement pagination; it loads all employee list items at once.
- With around 200 users, load time is nearly instant.
- Larger environments may benefit from implementing pagination or list filtering to improve performance.
