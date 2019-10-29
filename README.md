# Graph-API-O365-Reports-to-SQL
 
* Prerequisites: 
  * Module: Logging - To install "Install-Module Logging -verbose -Force"
  * Module: SQLSever - To install "Install-Module SqlServer -verbose -Force"
  * Azure App:
    * Application ID and Secret key
    * Permissions to Graph API Reporting

  * Use the Create-Database script file to create the:
    * Resource Group
    * SQL Server
    * Database
    * User
    * Firewall exceptions
        
     \* You can skip this part of you've created your DB already. 
    
  * Comment/Uncomment the Create-Table function at the bottom of the script when
        Creating or after creating the tables (respectively). You only need this one time. 

  * I have tested this on Windows 10 and Windows Server 2012 R2

  * Passwords and Secrets - prior to running the script, make sure to run the Create-Configuration
        first. The script will ask for all the necessary information and encrypt the passwords and
        secrets to a .json file. The .json file will picked by the main script for use. 

     \* The encrypted passwords can only be decrypted by the same user account who encrypted it
            when using the script on a different account, re-run the create-Configuration script.
        
