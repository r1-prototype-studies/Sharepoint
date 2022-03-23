<h1>Sharepoint</h1>

- [Notes](#notes)
- [Steps](#steps)
- [References](#references)

# Notes
* Creating self signed certificate is a one time activity. We need not create for all the web parts
* 
# Steps
1. Run the below command to create the sharepoint project in VS code
    ```
    yo @microsoft/sharepoint
2. Run the below command to create the self signed certificate 
    ```
    gulp trust-dev-cert
3. Open the file ./config/serve.json and change "enter-your-sharepoint-site" to the tenant we are using
4. Run the above command to start the application
    ```
    gulp serve
4. 

# References



