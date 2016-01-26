
var vm = {};

initModel = function(viewModel){
	vm = viewModel;
}

var Table = function(t){

    var self = this;
    self.name = t.name;
    self.localname = t.localname;
    self.timestamp = moment(t.timestamp).format("YYYY-MM-DD");
    self.invisible = t.invisible;
    self.guiFields = ko.observableArray();

    self.attributes = {};
    $.each(vm.tableAttributes, function(i, a){
        self.attributes[a] = t[a];
    });

    self.guiFields(ko.utils.arrayMap(ko.utils.arrayFilter(t.field,function(f){ return f.fieldtype != 255;}), function(f){
        return new Field(f, self.name);
    }));

    self.selected = ko.observable(false);
    self.shown = ko.computed(function(){
        return vm.shownTable() ? (vm.shownTable().name == self.name) : false; 
    });

    self.select = function(){
        self.selected(!self.selected());
    }
    self.show = function(){
        vm.shownTable(vm.shownTable() ? (vm.shownTable().name == self.name ? null: self) : self);
    }

    self.selectedFields = ko.computed(function(){
        return ko.utils.arrayFilter(self.guiFields(), function(f){
            return f.selected();
        });
    });

    self.filteredFields = ko.observableArray();

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
                if(item.timestamp.toLowerCase().indexOf(vm.fieldFilter().toLowerCase()) != -1){
                    return true;
                }
                return false;
            }));
        }else{  
            self.filteredFields(self.guiFields().slice());
        }
    }

    self.selectFields = ko.observable(false);


    self.selectFields.subscribe(function(newValue){
        ko.utils.arrayForEach(self.filteredFields(),function(item){
            item.selected(newValue);
        });
        self.selected(newValue);
    });

    self.filterFields();
}

var Field = function(f, tablename){
    var self = this;

    self.table = tablename;
    self.name = f.name;
    self.timestamp = moment(f.timestamp).format("YYYY-MM-DD");
    self.localname = f.localname;
    self.fieldtype = f.fieldtype;
    self.selected = ko.observable(false);

    self.attributes = {};
    $.each(vm.fieldAttributes, function(i, a){
        self.attributes[a] = f[a];
    });

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
    self.select = function(){
        self.selected(!self.selected());
    }
}

var Option = function(o){
    var self = this;
    self.text = o;
    this.select = function(){
        vm.status(this.text);
    }
}