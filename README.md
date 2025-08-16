E-commerce Platform on Google Apps Script
This repository contains the code for a full-stack e-commerce application built entirely on Google Apps Script. It leverages Google Sheets as a database, providing a robust, serverless, and easily manageable backend. The platform includes separate user and admin web interfaces for a complete end-to-end e-commerce solution.

⚡️ Key Features

User Authentication System: The Index.html file works with Code.gs to provide a secure user login and signup system. User credentials and data, including name, email, password, and balance, are stored and managed in a dedicated Google Sheet.






Dynamic Product Catalog: The Shop.html page dynamically fetches and displays a list of products from a Google Sheet. Users can browse, search, and sort products by various criteria like price, name, and rating.



Admin Management Panel: The admin.html file provides a comprehensive dashboard for administrators. Admins can manage users, products, and announcements, as well as view and process user requests.





Transactional Logic: The Code.gs script handles critical business logic:


Recharge Requests: Users can submit recharge requests with an amount and transaction ID. Admins can then approve or reject these requests, automatically updating the user's balance upon approval.




Purchase Workflow: The system allows users to submit purchase requests. Upon admin approval, the user's balance is deducted, and the product's stock is automatically updated.




Integrated Communication: A built-in chat system allows users to communicate directly with the administrator. The chat history is saved to a Google Sheet, enabling persistent conversations.



Serverless Architecture: The entire application runs on Google's infrastructure, eliminating the need for a separate server or hosting service. All data is stored in Google Sheets, making it easy to manage and access.



📂 Repository Structure
Code.gs: The core backend logic written in Google Apps Script. It contains functions for data manipulation, user authentication, and business processes.




Index.html: The user-facing HTML file for login and signup.

Shop.html: The user-facing HTML file for the product catalog, recharge, and chat functionalities.

admin.html: The HTML file for the administrative dashboard, offering full control over the platform.

appsscript.json: The Apps Script manifest file (not explicitly provided in the source, but required for the project). It defines project settings and required scopes.

🚀 How to Set It Up
Create Google Sheets: Create separate Google Sheets to act as your database. You will need one for 

Users , 

Products , 

Announcements , and 

Orders.

Populate Sheets with Headers: Add the predefined headers to each sheet as specified in Code.gs:


USERS: ['id', 'name', 'email', 'password', 'createdAt', 'lastLogin', 'balance', 'isActive', 'pending', 'notes'] 


PRODUCTS: ['id', 'name', 'description', 'price', 'discountPrice', 'imageUrl', 'rating', 'tags', 'category', 'stock', 'createdAt', 'updatedAt'] 


ANNOUNCEMENTS: ['id', 'title', 'content', 'createdAt', 'updatedAt'] 


ORDERS: ['id', 'userEmail', 'orderDate', 'productName', 'quantity', 'price', 'totalAmount', 'status'] 

Additionally, create sheets named 

PendingRecharges, PurchaseRequests, and ChatMessages within the USERS spreadsheet, with their respective headers.


Deploy as a Web App:

Copy the code from Code.gs, Index.html, Shop.html, and admin.html into a new Google Apps Script project.

In 

Code.gs, replace the placeholder IDs with the actual IDs of your Google Sheets.

Go to Deploy > New Deployment and select Web app.

Configure the deployment to allow access for "Anyone" or "Anyone, even anonymous".

Publish the app and note the URL provided.

Update Links: Update the SHOP_PAGE_URL in Index.html to the URL of your deployed web app.

Start Using: Open the web app URL in your browser to access the user interface. Access the admin panel by adding the query parameter ?admin=true to the web app URL.
