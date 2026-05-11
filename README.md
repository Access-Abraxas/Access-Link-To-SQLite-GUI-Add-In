# The Microsoft Access Link To SQLite Table GUI Add-In

This 100% Free, Open Source Microsoft Access Link-to-SQLite GUI Add-In tool was designed to help create MS Access linked tables to any SQLite database file.  This tool provides a GUI for connecting to the SQLite database and allows the user to select exactly which SQLite tables to link to your Access database.  This tool can link to tables in .sqlite, .sqlite3, .db, .db3, .s3db, and .sl3 database file types. The tool creates the connection string to SQLite, creates the linked-tables, and then establishes the connection to those SQLite tables in you Access database.  Once the SQLite tables are linked in your Access database, you will see them in the NavPane as normal linked tables, and you will be able to add, query, edit, and delete records in your SQLite database directly from your Microsoft Access database application.  

*** NOTE ***

Using the Access Link-to-SQLite Add-In tool requires a ODBC driver for SQLite installed on your system.  See the "Where to Download" below for information about where to get a SQLite ODBC driver for Windows, if you don't have one already. 

## Screenshots of the Access Link-to-SQLite GUI Add-In Tool:

A screenshot of the main form used within the Microsoft Access Link-to-SQLite GUI Add-In tool:

![Screenshot of the Microsoft Access Link-to-SQLite GUI Add-In form](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/blob/main/Screenshots/Access_Link_To_SQLite_GUI_Add-In_Form.jpg)



## Where to Download:

Download the FREE Microsoft Access Link-to-SQLite GUI Add-In (.accda file) here: [Latest ACCDA Download](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/raw/refs/heads/main/ACCDA/SqliteConnector.accda)

Or download the latest Stable Release package here: [Latest Full Release Package](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/releases)

If you don't have an SQLite ODBC driver already, you can download one [from here](http://www.ch-werner.de/sqliteodbc/)

- The 64-Bit SQLite ODBC driver can be [downloaded here](http://www.ch-werner.de/sqliteodbc/sqliteodbc_w64.exe)
- The 32-Bit SQLite ODBC driver can be [downloaded here](http://www.ch-werner.de/sqliteodbc/sqliteodbc.exe)

NOTE: ODBC drivers should match the same architecture as Microsoft Access.  So if you have 64-bit Access, use the 64-bit ODBC driver.



## System Requirements:

The following software is required to be installed to use this Microsoft Access Add-In tool:

1. Microsoft Access 2007/2010/2013/2016/2019/2021/2024/365 or Higher [More Info](https://www.microsoft.com/en-us/microsoft-365/access)
2. SQLite Database System [More Info](https://sqlite.org/)
3. An ODBC Driver for SQLite [More Info](http://www.ch-werner.de/sqliteodbc/)

NOTE: ODBC drivers should match the same architecture as Microsoft Access.  So if you have 64-bit Access, use the 64-bit ODBC driver.



## How to Install:

To just install the latest version of this Access Add-In without messing with any of the source code, complete the following steps:

1. Download the latest ACCDA file [from HERE](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/raw/refs/heads/main/ACCDA/SqliteConnector.accda).
2. Open any database in Microsoft Access and go to the "Database Tools" ribbon menu.

![Screenshot of the Microsoft Access Database Tools Ribbon Menu](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/blob/main/Screenshots/Microsoft_Access_Database_Tools_Ribbon_menu.jpg)
 
3. On the "Database Tools" ribbon, click the "Add-Ins" option drop down, and choose the "Add-In Manager" option.  This will open the Access "Add-In Manager" form.

![Microsoft Access Add-Ins Manager Menu](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/blob/main/Screenshots/Microsoft_Access_Add-Ins_menu.jpg)

4. The Access "Add-In Manager" form will be opened, click the "Add New..." button. 

![Microsoft Access Add-In Manager Add New Button](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/blob/main/Screenshots/Microsoft_Access_Add-In_Manager_form.jpg)

5. The Access "Open" file dialog will be opened. 

![Microsoft Access Open Add-In ACCDA file dialog](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/blob/main/Screenshots/Microsoft_Access_Open_Add-In_ACCDA_file.jpg)

6. On the "Open" file dialog, navigate to the ACCDA file you downloaded and click the "Open" button.  

![Open SqliteConnector.accda Add-In File in Microsoft Access](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/blob/main/Screenshots/Microsoft_Access_Open_SqliteConnector_accda_file.jpg)

7. The Add-In will be installed and will now show in the "Add-In Manager" form.

![Microsoft Access Add-In Manager with new Add-In](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/blob/main/Screenshots/Microsoft_Access_Add-In_Manager_with_New_Add-In.jpg)

The Microsoft Access Link-to-SQLite GUI Add-In should now be installed correctly.



## How to Use this Tool:

To use this tool once it has been installed:

1. Open any Access database in Microsoft Access.
2. On the "Database Tools" ribbon, click the "Add-Ins" option drop down, and choose the "Link to SQLite Database" option.  

![Microsoft Access Add-Ins Manager Menu](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/blob/main/Screenshots/Microsoft_Access_Add-Ins_Link_Sqlite_menu.jpg)

3. The "Access Link to SQLite" GUI form will open.

![The Microsoft Access Link-to-SQLite Add-In form](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/blob/main/Screenshots/Access_Link_To_SQLite_GUI_Add-In_Form.jpg)



## Project Resources:

Here is a list of some resources available for this project:

1. Log bugs and issues in our [Issues Database Here](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/issues)

2. Discussions about this project can be created in our [Discussion Forum Here](https://github.com/Access-Abraxas/Access-Link-To-SQLite-GUI-Add-In/discussions)



## Project Contributors and Credits:

A **GREAT BIG THANKS** to the following contributors to this project:

1. [The Microsoft Access Team](https://www.microsoft.com/en-us/microsoft-365/access) - A HUGE THANKS for creating the most AMAZING and versatile database product ever to exist, it has changed my life FOREVER!!!

2. [The SQLite Team](https://sqlite.org/crew.html) - Another HUGE THANKS for creating such an AMAZING database system too, a system which is up and coming in this brave new world and I'm loving it!

3. [Christian Werner](http://www.ch-werner.de/) - Thank you for the GREAT SQLite ODBC drivers for Windows!

4. [Albert Kallal](http://www.kallal.ca) - Thanks for inspiring me to use SQLite databases more!

5. [Crystal Long](https://msaccessgurus.com/) - Thanks for always being my biggest fan, especially when it comes to my work involving Microsoft Access!

6. [George Hepworth](https://www.gpcdata.com/) - Thanks for always hosting the Users Group Meetings and for coming up with GREAT ideas for Access database applications.

7. [The Access Users Groups](https://accessusergroups.org/) - A HUGE THANKS to the Access Pacific and Access Express Australia groups for inspiring me to use Access more and for always giving me ideas for database apps to create! 

8. [The AccGPT Tool for MS Access](https://accgpt.net) - For helping me build this tool in less than one day!

9. [The Access Add-In Helper Tool](https://github.com/Access-Abraxas/Access-Add-In-Helper) - For helping me create this Add-In from my initial Link-to-SQLite Access App in just a few minutes!
 
10. [Geoffrey Griffith](https://imaginethought.com/site/) - For my work to create this 100% FREE tool to help people connect Microsoft Access to their SQLite data.



## Other Free Microsoft Access Tools:

Some other projects I've created in MS Access that I've also made available 100% FREE:

1. [Access Add-In Helper](https://github.com/Access-Abraxas/Access-Add-In-Helper)

2. [Access Database Property Editor Add-In](https://github.com/Access-Abraxas/Access-Database-Property-Editor-Addin)

3. [Remove VBA Line Numbers Add-In](https://github.com/Access-Abraxas/Remove-VBA-Line-Numbers-Addin)

4. [Win32 API Declarations for VBA](https://github.com/Access-Abraxas/Win32-API-Declarations-for-VBA) 
