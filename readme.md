#LIP - Package Management for LIME Pro

LIP is a package management tool for LIME Pro. A package can currently contain declarations for fields and tables, VBA modules, localizations, LIME Bootstrap Apps and SQL-procedures. LIP downloads and installs packages from Package Stores. A Package Store is any valid source which serves correct JSON-files and package.zip files. Currently the LIME Bootstrap AppStore is the only available Package Store.

LIP is inspired from Pythons PIP and Nodes NPM but adapted for LIME Pro.

##Using LIP
The current implementation is written in VBA and is used in the immediate window in LIME Pro's VBA-editor.

###Get started
Before you can start installing packages, you need to install LIP to your LIME Pro database. Download the zip-file, which includes all necessary files to get started, and follow these steps:

1. Add the SQL-procedures to your database
2. Run `EXEC lsp_refreshldc` on your database
3. Import the `vba/lip.bas`-file to your VBA. Compile and save.

The first time you install a package, some necessary VBA-modules will automatically be installed.

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
A Package Store could either be file based or web based. A store has a fixed URL (example "http://limebootstrap.lundalogik.se/api/apps"). The URL has subdirectories for each app (example "./checklist"). If the source is a file-based a `/package.json` should automatically be append. This URL should return a JSON as:

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
                "name": "CSP_Test"
            }
        ],
        "tables": [
            {
                "name": "test",
                "descriptive":"[test].[title]",
                "localname_singular": 
                {
                    "sv": "Test",
                    "en-us": "Test"
                },
                "localname_plural": 
                {
                    "sv": "Test",
                    "en-us": "Test"
                },
                "attributes": {
                    "invisible": "no",
                    "actionpad":"lbs.html",
                    "policy":"policy_database_name"
                },
                "fields": [
                    {
                        "name": "title",
                        "type": "text",
                        "localname": {
                            "sv": "Titel",
                            "en-us": "Title"
                        }, 
                        "attributes": {
                            "length": 256
                        }
                    },
                    {
                        "name": "industry",
                        "type": "option",
                        "localname": {
                            "sv": "Titel",
                            "en-us": "Title"
                        },
                        "options":[
                            {
                                "key":"science",
                                "sv":"Vetenskap",
                                "en-us": "Science"
                            },
                            {
                                "key":"steel",
                                "sv":"Stål",
                                "en-us": "Steel",
                                "default":true
                            }
                        ],
                        "attributes": {
                            "length": 256
                        }
                    }
                ]
            }
        ]
    }
}
```

The installer should first see if a package is locally installed or not. If the package is installed local

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




