packagebuilder = {
    vm: {},
    initialize: function(vm){
        this.vm = vm;
    },
    "serializePackage": function(){
        // Serialize selected tables and fields and combine with localization data
            var data = {};
            var packageTables = [];
            var tables = [];
            var packageRelations = [];
            var relations = {};
            var sqlObjects = [];
            if (vm.name() == ""){
                alert("Package name is required");
                return;
            }
            try{
                // For each selected table
                
                $.each(vm.selectedTables(),function(i,table){
                    var packageTable = {};
                    //Clone the table object
                    packageTable = jQuery.extend(true,{},table);
                    // Fetch local names from table with same name
                    var localNameTable  = vm.localNames.Tables.filter(function(t){
                        return t.name == table.name;
                    })[0];

                    // Set singular and plural local names for table
                    packageTable.localname_singular = localNameTable.localname_singular;
                    packageTable.localname_plural = localNameTable.localname_plural;
                    
                    // For each selected field in current table
                    var fields = [];
                    var packageFields = [];
                    
                    var selectedFields = jQuery.extend(true,{},table.selectedFields());
                    $.each(selectedFields,function(j,field){
                        // Fetch local names from field with same name
                        var localNameField = localNameTable.Fields.filter(function(f){
                            return f.name == field.name;
                        })[0];
                        //Clone the field
                        var packageField = jQuery.extend(true,{},field);
                        
                        // Set local names for current field
                        packageField.localname = localNameField;
                        
                        //create relations
                        try{
                            if(field.attributes.fieldtype == "relation"){
                                //Lookup if relation already added
                                var existingRelation = relations[field.attributes.idrelation];
                                
                                if(existingRelation == null || existingRelation == undefined){
                                    var packageRelation = new Relation(field.attributes.idrelation,table.name, field.name);
                                    relations[field.attributes.idrelation] = packageRelation;
                                    
                                    
                                }
                                else{
                                    existingRelation.table2 = table.name;
                                    existingRelation.field2 = field.name;
                                }
                            }
                        }
                        catch(e){
                            alert(e);
                        }
                        
                        if(packageField.localname && packageField.localname.name){
                            delete packageField.localname.name;
                        }

                        if(packageField.localname && packageField.localname.order){
                            delete packageField.localname.order;
                        }
                        
                        //The separator is added correctly as a property on a field, instead of localname.
                        if(packageField.localname && packageField.localname.separator){
                            packageField.separator = packageField.localname.separator;
                            
                            delete packageField.localname.separator;
                                
                        }
                        
                        if(packageField.separator && packageField.separator.order){
                            delete packageField.separator.order;   
                        }
                        
                        if(packageField.localname && packageField.localname.option){
                            delete packageField.localname.option;
                        }

                        // Push field to fields
                        fields.push(packageField);
                        
                        
                    });
                    
                    
                   
                    
                    // Set fields to the current table
                    packageTable.fields = fields;
                    
                    // Push table to tables
                    packageTables.push(packageTable);
                });
                
                 
                //Add relations as the package expects
                for(idrelation in relations){
                    if(relations[idrelation].table2 != ""){
                        packageRelations.push({"table1": relations[idrelation].table1,
                                                "field1": relations[idrelation].field1,
                                                "table2": relations[idrelation].table2,
                                                "field2": relations[idrelation].field2
                                                })
                    }
                    
                }
                
                var packageRelationFields = [];
                //Fetch all relationfields in package
                var index;
                for(index = 0;index < packageTables.length; ++index){
                    var j;
                    for (j = 0;j <  packageTables[index].fields.length; j++){
                      var f = packageTables[index].fields[j];
                      if (f.attributes.fieldtype == "relation"){
                        packageRelationFields.push({ "name":packageTables[index].name + '.' + f.name, "remove": 1});   
                      }
                    }
                }
                
                //Check if field is existing in an relation
                for (index = 0;index < packageRelationFields.length; index++){
                    var rf = packageRelationFields[index];
                    var j;
                    for (j = 0; j < packageRelations.length;j++){
                        var rel = packageRelations[j];
                        if (rel.table1 + '.' + rel.field1 == rf.name || rel.table2 + '.' + rel.field2 == rf.name){
                            rf.remove = 0;
                        }
                    }
                }
                
                //remove unpaired relationfields 
                $.each(packageRelationFields,function(i,relField){
                    if(relField.remove == 1){
                        $.each(packageTables, function(j,packageTable){
                            if(packageTable.name == relField.name.substring(0, relField.name.indexOf("."))){
                                var indexOfObjectToRemove;
                                //find the field to remove
                                $.each(packageTable.fields, function(k, packageField){
                                    if (packageField.name == relField.name.substring(relField.name.indexOf(".") + 1)){
                                        indexOfObjectToRemove = k;
                                    }
                                });
                                //remove field from package
                                if(indexOfObjectToRemove){
                                    packageTable.fields.splice(indexOfObjectToRemove,1);
                                }
                            }
                        
                        
                    });
                }
                });
                
                $.each(vm.selectedSql(),function(i, sql){
                    sqlObjects.push({"name": sql.name, "definition": vm.sqlDefinitions()[sql.name]})
                });
                
                
                
            }
            catch(e){
                alert(e);
            }
            
            try {
                arrComponents = [];
                $.each(vm.selectedVbaComponents(), function(i, component){
                    arrComponents.push({"name": component.name, "relPath": "Install\\" + component.name + component.extension() })
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
                        "tables" : packageTables,
                        "vba" : arrComponents,
                        "relations": packageRelations,
                        "sql": sqlObjects
                        
                    }
                }
                //lbs.log.debug(JSON.stringify(data));
            }catch(e) {alert("Error serializing LIP Package:\n\n" + e);}
            
            // Save using VBA Method
            try{
                //Base64 encode the entire string, commas don't do well in VBA calls.
                lbs.common.executeVba('LIPPackageBuilder.CreatePackage', window.btoa(JSON.stringify(data)));
            }catch(e){alert(e);}
            
        
    }
}