# Sharepoint Service to Service Authorization. Integration with Business Central. No user permissions.
I think integrating Sharepoint into Business Central is a fairly common request. The standard library already supports some level of Sharepoint/OneDrive integration out of the box, but it is quite limited. What if we need more advanced logic? In that case, we can use Sharepoint Interfaces from the System Application! But did you know that this will only work if the user has permissions on Sharepoint? Regardless of whether we are using Microsoft Entra Application, it all comes down to the authorization that Sharepoint Interfaces use by default. Let's dive in and understand this better.

https://vld-nav.com/sharepoint-custom-authorization

![image](https://github.com/Drakonian/sharepoint-custom-authorization/assets/16802407/11d0497b-eff5-4307-8966-efedd5154d2e)


---

# How to Install a Per Tenant Extension (.app file) in Business Central

This guide provides step-by-step instructions for installing a Per Tenant Extension (PTE) `.app` file in Microsoft Dynamics 365 Business Central via the Extension Management page.

## Prerequisites

- **Permissions**: Ensure you have the permissions to manage extensions in Business Central.
- **.app File**: Have the `.app` file you wish to install ready.

## Installation Steps

### 1. Access the Extension Management Page

- Go to **Extension Management** from the Search.

### 2. Upload the Extension

- Click **Upload Extension** on the Extension Management page.
- Press **Select .app file DrillDown**, navigate to your `.app` file, select it, and click **Open**.
  
### 3. Install the Extension

- Click **Deploy**
- Accept any terms and conditions, if prompted.
- Confirm the installation by clicking **Yes**.

### 4. Verify Installation

- The installation process may take a few minutes. You can monitor the progress on the Extension Management page, where the status will change to **Installed** once completed.

## Troubleshooting

- **Installation Errors**: Refer to the error message details and consult the extension's documentation or support resources.
- **Permissions Issues**: Confirm you have the necessary permissions to install extensions. Contact your system administrator if you're unsure.


---
