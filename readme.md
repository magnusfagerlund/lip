#LIP - Package Management for LIME Pro

LIP is a package management tool for LIME Pro. A package can currently contain declerations for fields and tables, VBA modules, localizations, LIME Bootstrap Apps and SQL-procedures. LIP downloads and installs packages from Package Stores. A Package Store is any vaild source which serves correct JSON-files and package.zip files. Currently  the LIME Bootstrap AppStore is the only availible Package Store.

##Using LIP
The current implementation is written i VBA and is used in the intermediate window in LIME Pros VBA-editor. Simply import the `vba/lip.bas`-file to get started

###Install a package 
To install a package simply run
`lip.Install "ExamplePackage"`

###Install all dependencies for a LIME Pro solution
All installed packages are kept tracked of inside the `package.json`-file in the ActionPad folder. If you transfer this file to a new LIME Pro database you can use this file to conduct a brand new install. Just type
`lip.Install`

###Update a package
If a package already exist and should be updated or reinstalled you must explicitly use the update command 
`lip.Update "ExamplePackage"`

###Remove a package
Not yet implemented!

##Behind the scene
-