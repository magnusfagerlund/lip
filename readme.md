#LIP - Package Management for LIME Pro

LIP is a package management tool for LIME Pro. A package can currently contain declarations for fields and tables, VBA modules, localizations, LIME Bootstrap Apps and SQL-procedures. LIP downloads and installs packages from Package Stores. A Package Store is any valid source which serves correct JSON-files and package.zip files. You can configure which stores to use in the packages.json-file in your actionpad-folder.

LIP is inspired from Pythons PIP and Nodes NPM but adapted for LIME Pro.

##Using LIP
The current implementation is written in VBA and is used in the immediate window in LIME Pro's VBA-editor.

###Get started
Before you can start installing packages, you need to install LIP to your LIME Pro database. Download the zip-file, which includes all necessary files to get started, and follow these steps:

1. Add the SQL-procedures to your database by running the SQL-scripts (Important! Make sure you run the scripts on your database and NOT the master-database by selecting the correct database in the upper left corner)
2. Run `exec lsp_refreshldc` on your database
3. Import the `vba/lip.bas`-file to your VBA. Compile and save.

The first time you install a package, all necessary VBA-modules will automatically be installed.

###Install a package 
To install a package, simply type your command in the Immediate-window of the VBA. There are three different installation methods:

`lip.Install("Packagename")`
Standard installation of package. Searches for the package on the stores configured in your `packages.json`-file in the actionpad-folder.

`lip.InstallFromZip("SearchPathToZipFile")`
Install a package from a zip-file, provide the searchpath as an argument (including the filename).

`lip.InstallFromPackageFile`
All installed packages are kept tracked of inside the `packages.json`-file in the ActionPad folder. If you transfer this file to a new LIME Pro database you can use this file to conduct a brand new install by typing the command above.

###Update a package
If a package already exist and should be updated or reinstalled you must explicitly use the update command 

`lip.UpgradePackage("ExamplePackage")`

###Remove a package
__Not yet implemented!__
Should remove a specified package

`lip.Remove "ExamplePackage"`

###Freeze a solution to a package
__Not yet implemented!__
Compare to `pip freeze > requirement.txt`. Creates a package from a LIME Pro solution.

`lip.Freeze`

##Behind the scene



### A Package
A package is a ZIP-file containing all required resources to install a package

### A Package Store
A Package Store could either be file based or web based. A store has a fixed URL (example "http://limebootstrap.lundalogik.se/api/apps"). The URL has subdirectories for each app (example "./checklist"). If the source is a file-based a `/app.json` should automatically be append.

### app.json

An example of how the app.json-file could look like:

```JSON
{
    "name": "[NAME OF PACKAGE]",
    "author":"[AUTHORS NAME]",
    "status":"[STATUS OF THE PACKAGE, CAN BE: 'release', 'beta' OR 'Development']",
    "shortDesc":"[A short text to describe the package]",
    "versions":[
            {
            "version":"1",
            "date":"2014-02-06",
            "comments":"Css improvements!"
        },
        {
            "version":"0.9",
            "date":"2013-11-18",
            "comments":"The first stable beta of the Business Funnel"
        }
    ],
    "dependencies":{
        "vba_json":"1.0",
        "lime_basic":"5.0"
    },
    "install": {
        "localize": [
            {
                "owner": "checklist",
                "context": "title",
                "sv": "test",
                "en-us": "test",
                "no": "test",
                "fi": "test"
            }
        ],
        "vba": [
            {
                "relPath": "Install/Test.bas",
                "name": "Checklist"
            }
        ],
        "sql":[
        	{
                "relPath": "test.sql",
                "name": "csp_mypackage_test"
            }
        ],
        "tables": [
            {
                "name": "test",
                "localname_singular": 
                {
                    "sv": "Test",
                    "en_us": "Test"
                },
                "localname_plural": 
                {
                    "sv": "Test",
                    "en_us": "Test"
                },
                "attributes": {
                    "invisible": "2",
                    "descriptive":"[test].[title]",
                },
                "fields": [
                    {
                        "name": "title",
                        "localname": {
                            "sv": "Titel",
                            "en_us": "Title"
                        }, 
                        "attributes": {
                            "type": "text",
                            "length": 256
                        },
                        "separator": {
                            "sv": "Testseparator",
                            "en_us": "Test separator"
                        }
                    }
                ]
            }
        ]
    }
}
```

####localize

####sql
With this node you can add SQL-procedures and -functions. Specify the relative path to the SQL-file to run and the name of the procedure/function (used to check if it already exists and therefore should be updated instead of created). The name should include your packagename (e.g. csp_mypackage_addGoals) to distinguish the procedure from others. Don't forget to include the SQL-file in your package.

IMPORTANT: in your SQL-file you CAN'T include the rows before the "CREATE PROCEDURE"-row. The reason for this is how the procedure is added to the database. One SQL-file can also only include one procedure. Example of an SQL-file:
```
CREATE PROCEDURE [dbo].[csp_mypackage_addGoals]
    @@myVariable INT = NULL
AS
BEGIN
    SELECT 0
    --My code here
END

####vba
Here you can specify VBA-modules that should be installed. Please note that the VBA-file MUST be included in the zip-file of your package. Please specify the relativt path to the file and the name of the VBA-module.
####tables

#####name
Database name of the table. Example:
```
"name": "goaltable"
```
#####localname_singular
Localnames in singular. Each line in this node should represent one language. Valid languages are all languages LIME Pro supports. Example:
```
"localname_singular": {
"sv": "Måltabell",
"en_us": "Goal table"
}
```

#####localname_plural
Localnames in plural. Each line in this node represent a language. Valid languages are all languages LIME Pro supports. Example:
```
"localname_singular": {
"sv": "Måltabeller",
"en_us": "Goal tables"
}
```

#####attributes
Sets attributes for the table. Each line in this node represent an attribute. Valid attributes at the moment are:
tableorder, descriptive, invisible

#####fields

######name
The database name of the field. Example:
```
"name": "customernbr"
```

######localname
Localnames for the field. Each line in this node represent a language. Valid languages are all languages LIME Pro supports. Example:
```
"localname": {
"sv": "Kundnummer",
"en_us": "Customer number"
}
```

######separator
Adds a separator to the field. The separator is placed BEFORE the field. Specify the localnames for the separator inside this node. Example
```
"separator": {
"sv": "Administrativ information",
"en_us": "Administrative information"
}
```

######attributes
Sets attributes for the field. Each line in this node represent an attribute. Valid attributes at the moment are:
type, limereadonly, invisible, required, width, height, length, defaultvalue, limedefaultvalue, limerequiredforedit, newline, sql, onsqlupdate, onsqlinsert, fieldorder, isnullable.

### Versioning
####Package versioning
Packages should adhere to semantic versioning, example `1.0.0` or `MAJOR.MINOR.PATCH`. Please read [this](http://semver.org). 

Simplified:
`MAJOR`: Breaks backwards compatibility
`MINOR`: Adds new features but backward compatible
`PATCH`: Bugfixes

Minor and Patchs should always be upgrade to automatically if a dependency require it.

Major versions can only be upgraded to if explicit Upgrade command is used

####Dependency versioning
Stateing dependency verisons should adhere to [NPMs versioning](https://github.com/npm/node-semver)

A `version range` is a set of `comparators` which specify versions
that satisfy the range.

A `comparator` is composed of an `operator` and a `version`.  The set
of primitive `operators` is:

* `<` Less than
* `<=` Less than or equal to
* `>` Greater than
* `>=` Greater than or equal to
* `=` Equal.  If no operator is specified, then equality is assumed,
  so this operator is optional, but MAY be included.

For example, the comparator `>=1.2.7` would match the versions
`1.2.7`, `1.2.8`, `2.5.3`, and `1.3.9`, but not the versions `1.2.6`
or `1.1.0`.




