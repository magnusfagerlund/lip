
var vm = {};

// Load viewmodel
initModel = function(viewModel){
    vm = viewModel;
}

// Table object
var Table = function(t){

    var self = this;
    // Load database name
    self.name = t.name;
    // Load local name
    self.localname = t.localname;
    // Load timestamp
    self.timestamp = ko.observable(moment(t.timestamp).format("YYYY-MM-DD"));
    // Load invisible attribute
    self.invisible = t.invisible;
    // Initiate fields visible in gui
    self.guiFields = ko.observableArray();

    // Load attributes 
    self.attributes = {};
    $.each(vm.tableAttributes, function(i, a){
        self.attributes[a] = t[a];
    });

    // Load gui fields
    self.guiFields(ko.utils.arrayMap(ko.utils.arrayFilter(t.field,function(f){ return f.fieldtype != 255;}), function(f){
        return new Field(f, self.name);
    }));

    // If table is selected
    self.selected = ko.observable(false);

    // If fields of table are shown in column to the right
    self.shown = ko.computed(function(){
        return vm.shownTable() ? (vm.shownTable().name == self.name) : false; 
    });

    // Click function to select table
    self.select = function(){
        self.selected(!self.selected());
    };

    // Click function to show fields
    self.show = function(){
        vm.shownTable(vm.shownTable() ? (vm.shownTable().name == self.name ? null: self) : self);
    };

    // Computed for keeping track of selected fields
    self.selectedFields = ko.computed(function(){
        return ko.utils.arrayFilter(self.guiFields(), function(f){
            return f.selected();
        });
    });

    // Filtere fields. These are the ones shown in the gui (but based on guiFields)
    self.filteredFields = ko.observableArray();

    // Filter function for fields
    self.filterFields = function(){
        if(vm.fieldFilter() != ""){
            self.filteredFields.removeAll(); 
            self.filteredFields(ko.utils.arrayFilter(self.guiFields(), function(item) {

                if(item.name.toLowerCase().indexOf(vm.fieldFilter().toLowerCase()) != -1){
                    return true;
                }
                if(item.localname.toLowerCase().indexOf(vm.fieldFilter().toLowerCase()) != -1){
                    return true;
                }
                if(item.timestamp().toLowerCase().indexOf(vm.fieldFilter().toLowerCase()) != -1){
                    return true;
                }
                return false;
            }));
        }else{  
            self.filteredFields(self.guiFields().slice());
        }
    }

    // Select all fields
    self.selectFields = ko.observable(false);

    // Subscribe to select all event
    self.selectFields.subscribe(function(newValue){
        ko.utils.arrayForEach(self.filteredFields(),function(item){
            item.selected(newValue);
        });
        self.selected(newValue);
    });

    // Set default empty filter
    self.filterFields();
}

var Field = function(f, tablename){
    var self = this;

    // Field attributes
    self.table = tablename;
    self.name = f.name;
    self.timestamp = ko.observable(moment(f.timestamp).format("YYYY-MM-DD"));
    self.localname = f.localname;
    // self.fieldtype = vm.fieldTypes[f.fieldtype];
    
    
    self.attributes = {};
    $.each(vm.fieldAttributes, function(i, a){
        if(a == 'fieldtype'){
            self.attributes[a] = vm.fieldTypes[f[a]];
        }
        else if(a == 'relationsingle'){
            self.attributes["relationtab"] = f[a] == '0' ? '1':'0';
        }
        else{
            if(f[a]){
                self.attributes[a] = f[a];
            }
        }
    });

    // Observable for selecting field
    self.selected = ko.observable(false);

    // Subscribe to select event to see if table should be selected or deselected
    self.selected.subscribe(function(newValue){
        if(newValue){
            vm.shownTable().selected(newValue);
        }
        else{
            var checked = false;

            ko.utils.arrayForEach(vm.shownTable().guiFields(), function(item){
                checked = item.selected() ? true : checked;
            });
            vm.shownTable().selected(checked);
        }
    })
    // Click function for select
    self.select = function(){
        self.selected(!self.selected());
    }
}

// Status options (development, beta, release)
var StatusOption = function(o){
    var self = this;
    self.text = o;
    this.select = function(){
        vm.status(this.text);
    }
}

var VbaComponent = function(c){
    var self = this;
    self.name = c.name;
    self.type = c.type;
    self.selected = ko.observable(false);
    self.extension = function(){
        if(self.type == "Module"){
            return ".bas";
        }
        else if(self.type=="Class Module"){
            return ".cls";
        }
        else if(self.type == "Form"){
            return ".frm";
        }
        else return "";
    }
}

var Relation = function(idrelation, tablename, fieldname){
    var self = this;
    self.idrelation = idrelation;
    self.table1 =  tablename;
    self.field1 =  fieldname;
    self.table2 =  "";
    self.field2 = "";
    
    self.serialize = function(){
            return {    "table1": self.table1,
                        "field1": self.field1,
                        "table2": self.table2,
                        "field2": self.field2
                    };
    
    }
    
}

var SqlComponent = function(sql){
    var self = this;
    self.name = sql.name;
    self.selected = ko.observable(false);
}

