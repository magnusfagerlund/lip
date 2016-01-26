lbs.apploader.register('LIPPackageBuilder', function () {
    var self = this;

    /*Config (version 2.0)
        This is the setup of your app. Specify which data and resources that should loaded to set the enviroment of your app.
        App specific setup for your app to the config section here, i.e self.config.yourPropertiy:'foo'
        The variabels specified in "config:{}", when you initalize your app are available in in the object "appConfig".
    */
    self.config =  function(appConfig){
            this.yourPropertyDefinedWhenTheAppIsUsed = appConfig.yourProperty;
            this.dataSources = [];
            this.resources = {
                scripts: ['model.js'], // <= External libs for your apps. Must be a file
                styles: ['app.css'], // <= Load styling for the app.
                libs: ['json2xml.js'] // <= Allready included libs, put not loaded per default. Example json2xml.js
            };
    };

    //initialize
    /*Initialize
        Initialize happens after the data and recources are loaded but before the view is rendered.
        Here it is your job to implement the logic of your app, by attaching data and functions to 'viewModel' and then returning it
        The data you requested along with localization are delivered in the variable viewModel.
        You may make any modifications you please to it or replace is with a entirely new one before returning it.
        The returned viewModel will be used to build your app.
        
        Node is a reference to the HTML-node where the app is being initalized form. Frankly we do not know when you'll ever need it,
        but, well, here you have it.
    */
    self.initialize = function (node, vm) {

        $('title').html('LIP');


        vm.fieldTypes = {
            "1" : "string",
            "2" : "geography",
            "3" : "integer",
            "4" : "decimal",
            "7" : "time",
            "8" : "text",
            "9" : "script",
            "10" : "html",
            "11" : "xml",
            "12" : "link",
            "13" : "yesno",
            "14" : "multirelation",
            "15" : "file",
            "16" : "relation",
            "17" : "user",
            "18" : "security",
            "19" : "calendar",
            "20" : "set",
            "21" : "option",
            "22" : "image",
            "23" : "formatedstring",
            "25" : "automatic",
            "26" : "color",
            "27" : "sql",
            "255" : "system"
        }
        // Attributes for tables
        vm.tableAttributes = [
            "tableorder",
            "invisible",
            "descriptive",
            "syscomment",
            "label",
            "log",
            "actionpad"
        ];

        // Attributes for fields
        vm.fieldAttributes = [
            "fieldtype",
            "limereadonly",
            "invisible",
            "required",
            "width",
            "height",
            "length",
            "defaultvalue",
            "limedefaultvalue",
            "limerequiredforedit",
            "newline",
            "sql",
            "onsqlupdate",
            "onsqlinsert",
            "fieldorder",
            "isnullable",
            "type",
            "relationtab",
            "syscomment",
            "formatsql",
            "limevalidationrule",
            "label",
            "adlabel"
        ]
        
        // Checkbox to select all tables
        vm.selectTables = ko.observable(false);
        vm.selectTables.subscribe(function(newValue){
            ko.utils.arrayForEach(vm.filteredTables(),function(item){
                item.selected(newValue);
            });
        });

        // Navbar function to change tab
        vm.showTab = function(t){
            vm.activeTab(t);
        }

        // Set default tab to details
        vm.activeTab = ko.observable("details");


        // Serialize selected tables and fields and combine with localization data
        vm.serializePackage = function(){
            var data = {};
            var tables = [];

            // For each selected table
            $.each(vm.selectedTables(),function(i,table){
                
                // Fetch local names from table with same name
                var localNameTable  = vm.localNames.Tables.filter(function(t){
                    return t.name == table.name;
                })[0];

                // Set singular and plural local names for table
                table.localname_singular = localNameTable.localname_singular;
                table.localname_plural = localNameTable.localname_plural;
                
                // For each selected field in current table
                var fields = [];
                $.each(table.selectedFields(),function(j,field){
                    // Fetch local names from field with same name
                    var localNameField = localNameTable.Fields.filter(function(f){
                        return f.name == field.name;
                    })[0];
                    
                    // Set local names for current field
                    field.localname = localNameField;
                
                    if(field.localname && field.localname.name)
                        delete field.localname.name;

                    if(field.localname && field.localname.order)
                        delete field.localname.order;

                    if(field.separator && field.separator.order)
                        delete field.separator.order;   

                    if(field.localname && field.localname.option)
                        delete field.localname.option;

                    // Push field to fields
                    fields.push(field);
                });

                // Set fields to the current table
                table.fields = fields;

                // Push table to tables
                tables.push(table);
            });

            // Build package json from details and database structure
            data = {
                "name": vm.name(),
                "author": vm.author(),
                "status": vm.status(),
                "shortDesc": vm.description(),
                "versions":[
                    {
                    "version": vm.versionNumber(),
                    "date": moment().format("YYYY-MM-DD"),
                    "comments": vm.comment()
                }],
                "install" : {
                    "tables" : tables

                }
            }

            // Save to file using microsofts weird ass self developedd file saving stuff
            var blobObject = new Blob([JSON.stringify(data)]); 
            window.navigator.msSaveBlob(blobObject, 'package.json')
        }

        // Function to filter tables
        vm.filterTables = function(){
            if(vm.tableFilter() != ""){
                vm.filteredTables.removeAll(); 

                // Filter on the three visible columns (name, localname, timestamp)
                vm.filteredTables(ko.utils.arrayFilter(vm.tables(), function(item) {
                    if(item.name.toLowerCase().indexOf(vm.tableFilter().toLowerCase()) != -1){
                        return true;
                    }
                    if(item.localname.toLowerCase().indexOf(vm.tableFilter().toLowerCase()) != -1){
                        return true;
                    }
                    if(item.timestamp().toLowerCase().indexOf(vm.tableFilter().toLowerCase()) != -1){
                        return true;
                    }
                    return false;
                }));
            }else{  
                vm.filteredTables(vm.tables().slice());
            }
        }

        // Filter observables
        vm.tableFilter = ko.observable("");
        vm.fieldFilter = ko.observable("");

        // Load databas structure
        try{
            var db = {};
            lbs.loader.loadDataSource(db, { type: 'storedProcedure', source: 'csp_lip_getxmldatabase_wrapper', alias: 'structure' }, false);
            vm.datastructure = db.structure.data;
        }
        catch(err){
            alert(err)
        }
        // Data from details
        vm.author = ko.observable("");
        vm.comment = ko.observable("");
        vm.description = ko.observable("");
        vm.versionNumber = ko.observable("");
        vm.name = ko.observable("");
        // Set default status to development
        vm.status = ko.observable("Development");

        // Set status options 
        vm.statusOptions = ko.observableArray([
            new StatusOption('Development'), new StatusOption('Beta'), new StatusOption('Release')
        ]);
        
        // Load localization data
        try{
            var localData = {};
            lbs.loader.loadDataSource(localData, { type: 'storedProcedure', source: 'csp_lip_getlocalnames', alias: 'localNames' }, false);
            vm.localNames = localData.localNames.data;
        }
        catch(err){
            alert(err)
        }
        // Table for which fields are shown
        vm.shownTable = ko.observable();
        // All tables loaded
        vm.tables = ko.observableArray();
        // Filtered tables. These are the ones loaded into the view
        vm.filteredTables = ko.observableArray();
        // Load model objects
        initModel(vm);

        // Populate table objects
        vm.tables(ko.utils.arrayMap(vm.datastructure.table,function(t){
            return new Table(t);
        }));
  
        // Computed with all selected tables
        vm.selectedTables = ko.computed(function(){
            return ko.utils.arrayFilter(vm.tables(), function(t){
                return t.selected();
            });
        });

        // Subscribe to changes in filters
        vm.fieldFilter.subscribe(function(newValue){ 
            vm.shownTable().filterFields();
        })
        vm.tableFilter.subscribe(function(newValue){
            vm.filterTables();
        });
        
        // Set default filter
        vm.filterTables();

        return vm;
    };


    
});


ko.bindingHandlers.stopBubble = {
  init: function(element) {
    ko.utils.registerEventHandler(element, "click", function(event) {
         event.cancelBubble = true;
         if (event.stopPropagation) {
            event.stopPropagation(); 
         }
    });
  }
};