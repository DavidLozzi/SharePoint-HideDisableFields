"use strict"

/*
Created by David Lozzi, @davidlozzi, www.davidlozzi.com, 1/10/2014
Last Updated 7/5/2019
Requires jQuery, you can include: <script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
*/

const fields = {
    spContext: null,
    currentUser: null,
    peopleLoopCount: 0,
    peoplePickerBorder: "",
    metadataLoop: 0,

    getFieldType: (theCell) => {
        let fieldType = '';
        theCell.contents().each((i, obj) => {
            if (obj.nodeType === 8) {
                const ftRegx = new RegExp('(FieldType=")([a-zA-Z]*)"','gmi');
                const vals = ftRegx.exec(obj.data);
                if(vals && vals.length > 2 ) {
                    fieldType = vals[2];
                }
            }
        });
        return fieldType;
    },

    getRow: (fieldDisplayName) => {
        const theCell = fields.getCell(fieldDisplayName);
        const theRow = theCell.parent();
        return theRow;
    },

    getCell: (fieldDisplayName) => {
        const theCell = $(".ms-formbody").filter((i, o) => {
            const rx = 'FieldName=\"' + fieldDisplayName + '\"';
            const html = o.innerHTML.match(new RegExp(rx));
            return html !== null; // && html.length > 0;
        });
        return theCell;
    },
    
    disable: (fieldDisplayName) => {
        const theCell = fields.getCell(fieldDisplayName);
        const fieldType = fields.getFieldType(theCell);
        
        let value = "";
        let theControls;

        switch (fieldType) {
            case 'SPFieldLookup':
                fields.disableLookupField(theCell);
                break;
            case 'SPFieldMultiChoice':
                fields.disableMultiSelectField(theCell);
                break;
            case 'SPFieldTaxonomyFieldType':
                fields.disableMetadata(theCell);
                break;
            case 'SPFieldUser':
                fields.disablePeoplePicker(theCell);
                break;
            case 'SPFieldBoolean':
                fields.disableYesNo(theCell);
                break;
            case 'SPFieldChoice':
                if (theCell.find("input:checked").val() === "FillInButton" || theCell.find("input:checked").val() === "Specify your own value:") {
                    fields.disableOtherOption(theCell);
                } else if (theCell.find("input[type='radio']").length > 0) {
                    fields.disableRadioField(theCell);
                } else {
                    fields.defaultType(theCell);
                }
                break;
            default:
                fields.defaultType(theCell);
                break;
        }
    },

    defaultType: (theCell) => {
        const theControls = theCell.find("input,select,textarea,img");
        const value = "<span class='readonly'>" + theControls.val() + "</span>";
        
        theControls.hide();
        theCell.prepend(value);
        theCell.find("div.ms-inputBox").hide();
    },

    disableOtherOption: (theCell) => {
        const theControls = theCell.find("input,select,textarea,img,label,table");
        const fillInValue = theCell.find("input[id$='FillInChoice'], input[id$='FillInText']").val();
        const value = "<span class='readonly'>" + fillInValue + "<span>";
        theControls.hide();
        theCell.prepend(value);
    },

    disablePeoplePicker: (theCell) => {
        if ($(".sp-peoplepicker-topLevel", theCell).css("border") != "none") {
            //console.log("disablePeoplePicker")
            $(".sp-peoplepicker-editorInput", theCell).attr("disabled", "true")
            fields.peoplePickerBorder = $(".sp-peoplepicker-topLevel", theCell).css("border");
            $(".sp-peoplepicker-topLevel", theCell).css("border", "none");
            $(".sp-peoplepicker-delImage", theCell).hide();

            fields.peopleLoopCount++;
            //hate this loop, not sure how else to handle waiting for the objs 
            if (fields.peopleLoopCount < 3) {
                setTimeout(() => { fields.disablePeoplePicker(theCell) }, 350);
            } else {
                fields.peopleLoopCount = 0;
            }
        }
    },

    disableMetadata: (theCell) => {
        if ($(".valid-text", theCell).length > 0 || fields.metadataLoop == 5) {
            //console.log("found " + $(".valid-text",theCell).length + " with text " + $(".valid-text",theCell).text());
            const metaValue = $(".valid-text", theCell).text();
            const theControls = theCell.find("input,select,textarea,img");
            const value = "<span class='readonly'>" + metaValue + "<span>";
            theControls.hide();
            theCell.prepend(value);
            theCell.find("div.ms-inputBox").hide();
            fields.metadataLoop = 0;
        } else {
            fields.metadataLoop++;
            //yup, hate looping to wait for the DOM to load. Must. Find. Fix.
            if (fields.metadataLoop <= 5) {
                setTimeout(() => { fields.disableMetadata(theCell) }, 350);
            }
        }
    },

    disableLookupField: (theCell) => {
        const theControls = theCell.find("select");
        const selectedValue = theControls.find("option:selected");
        const value = "<span class='readonly'>" + selectedValue.text() + "<span>";
        theControls.hide();
        theCell.prepend(value);
    },

    disableMultiSelectField: (theCell) => {
        const theControls = theCell.find("table");
        const selectedValues = theControls.find("input:checked");
        let value = "";
        selectedValues.each((i, o) => {
            const selectedLabel = $(o).siblings("label");
            if (selectedLabel.text() === "Specify your own value:")
            {
                value += "<span class='readonly'>" + theCell.find("input[id$=FillInText]").val() + "</span><br/>";
            } else {
                value += "<span class='readonly'>" + selectedLabel.text() + "<span><br/>";
            }
        })
        theControls.hide();
        theCell.prepend(value);
    },
    
    disableYesNo: (theCell) => {
        const theControls = theCell.find("input");
        const isChecked = theControls.is(":checked");
        const value = "<span class='readonly'>" + (isChecked ? "Yes" : "No") + "</span>";
        theControls.hide();
        theCell.prepend(value);
    },

    disableRadioField: (theCell) => {
        const theControls = theCell.find("table");
        let selectedValue = theCell.find("input:checked").val();
        if(selectedValue === "DropDownButton") {
            selectedValue = theCell.find("option:selected").val();
        }
        const value = "<span class='readonly'>" + selectedValue + "</span><br/>";
        theControls.hide();
        theCell.prepend(value);
    },

    enable: (fieldDisplayName) => {
        const theCell = fields.getCell(fieldDisplayName);
        if (theCell.find("[class^='sp-peoplepicker']").length > 0) {
            fields.enablePeoplePicker(theCell);
        } else if (theCell.find(".ms-taxonomy-fieldeditor").length > 0) {
            fields.enableMetadata(theCell);
        } else {
            const theControls = theCell.find("input,select,textarea,img");
            const value = theCell.find("span.readonly")
            theControls.show();
            value.remove();
        }
    },

    enablePeoplePicker: (theCell) => {
        if ($(".sp-peoplepicker-topLevel", theCell).css("border") != fields.peoplePickerBorder) {
            //console.log("enablePeoplePicker")
            $(".sp-peoplepicker-editorInput", theCell).removeAttr("disabled")
            $(".sp-peoplepicker-topLevel", theCell).css("border", fields.peoplePickerBorder);
            $(".sp-peoplepicker-delImage", theCell).show();
            $("span.readonly", theCell).remove();
            fields.peopleLoopCount++;
            if (fields.peopleLoopCount < 7) {
                setTimeout(() => { fields.enablePeoplePicker(theCell) }, 350);
            } else {
                fields.peopleLoopCount = 0;
            }
        }
    },

    enableMetadata: (theCell) => {
        console.log("enableMetadata");
        if (fields.metadataLoop == 0) {
            //console.log("found " + $(".valid-text",theCell).length + " with text " + $(".valid-text",theCell).text());
            const theControls = theCell.find("input,select,textarea,img");
            theControls.show();
            $("div.ms-inputBox", theCell).show();
            $("span.readonly", theCell).remove();
        } else {
            setTimeout(() => { fields.enableMetadata(theCell) }, 350);
        }
    },

    hide: (fieldDisplayName) => {
        const theRow = fields.getRow(fieldDisplayName);
        theRow.hide();
    },

    show: (fieldDisplayName) => {
        const theRow = fields.getRow(fieldDisplayName);
        theRow.show();
    },

    //group names in an array of strings, i.e. ["Group One","Group Two"]
    disableWithAllowance: (fieldName, groups) => {
        fields.disable(fieldName);

        fields.spContext = new SP.ClientContext.get_current();
        fields.currentUser = fields.spContext.get_web().get_currentUser();

        fields.spContext.load(fields.currentUser);
        fields.spContext.load(fields.currentUser.get_groups());
        fields.spContext.executeQueryAsync(() => {
            fields.getGroupsAndEnable(fieldName, groups)
        }, asyncFailed);
    },

    getGroupsAndEnable: (fieldName, groups) => {
        let allowedToEdit = false;
        if (fields.currentUser.get_isSiteAdmin()) {
            allowedToEdit = true;
        } else {
            const groupEnum = fields.currentUser.get_groups().getEnumerator();
            while (groupEnum.moveNext()) {
                const group = groupEnum.get_current();
                if ($.inArray(group.get_title(), groups) > -1) {
                    allowedToEdit = true;
                    break;
                }
            }
        }
        if (allowedToEdit) {
            fields.enable(fieldName);
        }
    },

    //group names in an array of strings, i.e. ["Group One","Group Two"]
    hideWithAllowance: (fieldName, groups) => {
        fields.hide(fieldName);

        fields.spContext = new SP.ClientContext.get_current();
        fields.currentUser = fields.spContext.get_web().get_currentUser();

        fields.spContext.load(fields.currentUser);
        fields.spContext.load(fields.currentUser.get_groups());
        fields.spContext.executeQueryAsync(() => {
            fields.getGroupsAndShow(fieldName, groups);
        }, asyncFailed);
    },

    getGroupsAndShow: (fieldName, groups) => {
        let allowedToEdit = false;
        if (fields.currentUser.get_isSiteAdmin()) {
            allowedToEdit = true;
        } else {
            const groupEnum = fields.currentUser.get_groups().getEnumerator();
            while (groupEnum.moveNext()) {
                const group = groupEnum.get_current();
                if ($.inArray(group.get_title(), groups) > -1) {
                    allowedToEdit = true;
                    break;
                }
            }
        }
        if (allowedToEdit) {
            fields.show(fieldName);
        }
    },

    asyncFailed: (sender, args) => {
        alert('request failed ' + args.get_message() + '\n' + args.get_stackTrace());
    },

    setDefaultValue: (fieldDisplayName, stringValue) => {
        //TODO flush out the many field type options
        const theCell = fields.getCell(fieldDisplayName);
        const selectObj = theCell.find("SELECT");
        selectObj.val(stringValue);
    },
};

const Lozzi = window.Lozzi || {};
Lozzi.Fields = fields;