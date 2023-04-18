# Operate and Collect

Operate and Collect is a comprehensive solution designed to streamline workflow and improve data accuracy, provides a secure and reliable way to manage business operational data.

The solution comprises three parts:
1. Operate'n'Collect, a **Canvas Power App** business application that helps users manage revenue, cash-flow, business growth, profitability, and control costs, leading to more informed and timely decision-making.
2. **SharePoint Online** (site and several lists on the site) serves as a cloud-based database layer for operational data.
3. **Power BI report** is an analytics tool that provides businesses with valuable insights into their daily operations.

**The solution can be customized to meet the unique needs of each business**, feel free to contact us at contact@a3cloud.org

## Key Features 
* The managed solution provides the Canvas PowerApp that enables users to manage up to nine operational indicators per row of objects.
* Сapability to assign and manage access rights for various objects, brands or resources based on administrative user roles.
* Uses SharePoint Online (4 lists at least) as a database layer.
* Contains PowerBI report sample to visualize the data and get analytics.
* Application administrator can manage users directly from the application interface.
* Easy to use and customize UI.
* Scalable and flexible to meet changing business needs.
* Capability to integrate with other systems and applications to import-export data. 

## Technologies 
Operate and Collect solution was built using the following technologies: 
* Power Apps
* SharePoint Online
* PowerBI 

## Initial deployment

_An appropriate administrative role is required for the initial deployment. This particular role should be defined within the organization's case, whether it is a Global Administrator, Application Administrator, or another role that allows for such deployment within the tenant._

To fully install the Operate and Collect solution, you need to deploy and configure the application together with the database layer. Based on user experience feedback, we recommend creating a SharePoint Online site and SharePoint lists within it as the first step. Here is a PowerShell script prepared for this purpose.  
You can also deploy and configure the SharePoint lists according to the data layer description placed in the SharePoint online folder of the repository. Alternatively, you can connect to another data source with the same model structure for your deployment. Once the data layer is settled, you can import the application as a managed PowerPlatform solution and easily configure connections with data. The analytical report also works with the data structure from the data source, so it can be deployed as the last step of the initial process. Note that the initial deployment could be carried out separately for each part of the solution.

### Sharepoint lists  
Refer to Sharepoint folder for more info.
_PowerShell script to create site and 4 lists within your tenant_ 

### Operate'n'Collect application  
Refer to the app solution folder for more info. 
Clone the managed solution from the repository folder to your local location/repo. Import the solution into the environment within your tenant. Then, open the OnC application in Power Apps Studio to edit it. Configure the database connection in the Power Apps environment, and publish the application to your Power Apps environment. Finally, share the app with users.

### Analytics report
Reffer to PowerBI folder for more info.

## Usage
_Users should have proper licensing and administrative role to use PowerApps application._ 
The solution core provide users with functionality of adding new records, update, delete, use analytic report. Administrating functinons allows to manage objects/locations, provide users permissions to manage records for the objects.

## Limitations
The solution based on PowerPlatform, SharepointOnline, PowerBI, so inherits their limitations.

## Documentation 
User Guide 

## Roadmap 
Here are some of the planned updates for Operate and Collect: 
* Improved integration with third-party accounting software.
* Additional analytics and reporting features.
* Upgrading data security management options. 

## Feedback
We welcome feedback from users. If you have any suggestions or feature requests, any questions or issues with Operate and Collect, please contact our support team at https://a3cloud.org/support or via email contact@a3cloud.org

## License 
This project is licensed under the MIT License - see the LICENSE.md file for details.
