# AliexpressOrders
Now track your own aliexpress orders easily. This python script generates a json representation of all your orders.
It also saved the information in a google sheet if you have set up one. 

This will enable you to use powerful spreadsheet filters in google sheets for better order management and tracking.

### Usage
Get only screenshot:
`python aliexpress.py screenshot <screenshot_filename.png>`

Get only order details in json format:
`python aliexpress.py json`

Get both screenshot and order details:
`python aliexpress.py json,screenshot <screenshot_filename.png>`

### Whats working
1. Awaiting Shipments orders details
2. Awaiting Delivery Order details
3. Order Completed details
4. Orders pending payment details
3. Implement pagination for the above sections such that more than 10 orders per section appears
4. Implement google sheet integration
5. Implement batch google sheet update (Now update to google sheets is much faster - 20s for 200 plus orders)
6. Order tracking ID integrated. Also, the last package delivery status for the tracking id is retrieved
7. Scraps in background, tested with both chromium and firefox headless
8. Saves screenshot of the orders page in png format

### Work In Progress
1. Order carrier retrieval
2. Fix date and Purchase protection remaining data format for easier filtering in google sheets

### To Do
1. Integrate Tracking ID with an existing tracker to get package logistic updates

### Installation
* As of the current state, The package is dependent on lxml, pyquery, selenium and Chromedriver/PhantomJS Package
* Edit the path to Chromexdriver in the file
* Get Google Service Credentials. Download the credential json in same folder and point the path in the credentials call in file
* You need to share a google sheet and copy the url to an environment variable **AE_gsheet_url**
* Also, setup the Aliexpress Username and Password as environment variables, **AE_username** and **AE_passwd**

### License
This code ia available as free to use/redistribute under MIT License. Please check the LICENSE File for sharing and attribution requirements
