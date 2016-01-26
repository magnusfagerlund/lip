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


        vm.tableAttributes = [
            "tableorder",
            "invisible",
            "descriptive",
            "syscomment",
            "label",
            "log",
            "actionpad"
        ];

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
        $('title').html('LIP');


        vm.selectTables = ko.observable(false);
        vm.selectTables.subscribe(function(newValue){
            ko.utils.arrayForEach(vm.filteredTables(),function(item){
                item.selected(newValue);
            });
        });

        vm.showTab = function(t){
            vm.activeTab(t);
        }

        vm.serializePackage = function(){
            var data = {};
            var tables = [];

            $.each(vm.selectedTables(),function(i,table){
                
                var localNameTable  = vm.localNames.Tables.filter(function(t){
                    return t.name == table.name;
                })[0];

                table.localname_singular = localNameTable.localname_singular;
                table.localname_plural = localNameTable.localname_plural;
                
                var fields = [];
                $.each(table.selectedFields(),function(j,field){
                    var localNameField = localNameTable.Fields.filter(function(f){
                        return f.name == field.name;
                    })[0];
                
                    field.localname = localNameField;
                
                    if(field.localname && field.localname.name)
                        delete field.localname.name;

                    fields.push(field);
                });
                table.fields = fields;
                tables.push(table);
            });

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
            var blobObject = new Blob([JSON.stringify(data)]); 
            window.navigator.msSaveBlob(blobObject, (vm.name() || 'lippackage') + '.json')
        }

        vm.filterTables = function(){
            if(vm.tableFilter() != ""){
                vm.filteredTables.removeAll(); 
                vm.filteredTables(ko.utils.arrayFilter(vm.tables(), function(item) {

                    if(item.name.toLowerCase().indexOf(vm.tableFilter().toLowerCase()) != -1){
                        return true;
                    }
                    if(item.localname.toLowerCase().indexOf(vm.tableFilter().toLowerCase()) != -1){
                        return true;
                    }
                    if(item.timestamp.toLowerCase().indexOf(vm.tableFilter().toLowerCase()) != -1){
                        return true;
                    }
                    return false;
                }));
            }else{  
                vm.filteredTables(vm.tables().slice());
            }
        }

        var db = {};
        lbs.loader.loadDataSource(db, { type: 'storedProcedure', source: 'csp_lip_getxmldatabase_wrapper', alias: 'structure' }, false);
        vm.datastructure = db.structure.data;

        vm.activeTab = ko.observable("details");

        vm.author = ko.observable("");
        vm.comment = ko.observable("");
        vm.description = ko.observable("");
        vm.versionNumber = ko.observable("");
        vm.name = ko.observable("");
        vm.status = ko.observable("Development");

        vm.tableFilter = ko.observable("");
        vm.fieldFilter = ko.observable("");
        vm.shownTable = ko.observable();
        var localData = {};
        lbs.loader.loadDataSource(localData, { type: 'storedProcedure', source: 'csp_lip_getlocalnames', alias: 'localNames' }, false);

        vm.localNames = localData.localNames.data;
       
        vm.tables = ko.observableArray();
        vm.filteredTables = ko.observableArray();

        initModel(vm);

        vm.tables(ko.utils.arrayMap(vm.datastructure.table,function(t){
            return new Table(t);
        }));
  
        vm.selectedTables = ko.computed(function(){
            return ko.utils.arrayFilter(vm.tables(), function(t){
                return t.selected();
            });
        });

        vm.fieldFilter.subscribe(function(newValue){ 
            vm.shownTable().filterFields();
        })
        

        vm.statusOptions = ko.observableArray([
            new Option('Development'), new Option('Beta'), new Option('Release')
        ]);


        

        vm.tableFilter.subscribe(function(newValue){vm.filterTables()});

        

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