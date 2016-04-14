enums = {
    "initialize": function (vm){
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
        };
        
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
            "adlabel",
            "idrelation",
            "relationsingle"
        ];
        vm.FieldtTypeDisplayNames = {
            "string" : "Text",
            "formatedstring" : "Formatted text",
            "yesno" : "Yes/No",
            "link" : "Link",
            "option" : "Option",
            "relation" : "Realtion",
            "time" : "Time",
            "integer" : "Integer",
            "decimal" : "Decimal"
        };


    }
}