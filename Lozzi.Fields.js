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
        theCell.contents().map(() => {
            if (this.nodeType === 8) {
                console.log(this);
                const ftRegx = new RegExp('(FieldType=")([a-zA-Z]*)"','gmi');
                const vals = ftRegx.exec(this.data);
                if(vals && vals.length > 2 ) {
                    console.log(vals[2]);
                    return vals[2];
                }
            }
        });
    },
    
    disable: (fieldDisplayName) => {
        const theCell = fields.getCell(fieldDisplayName);
        const fieldType = fields.getFieldType(theCell);
        
        let value = "";
        let theControls;

        switch (fieldType) {
            case "":
                break;
            default:
        }

        if (theCell.find("[class^='sp-peoplepicker']").length > 0) {
            fields.disablePeoplePicker(theCell);
        } else if (theCell.find(".ms-taxonomy-fieldeditor").length > 0) {
            fields.disableMetadata(theCell);
        } else if (theCell.find("[id*='$LookupField']").length > 0) {
            fields.disableLookupField(theCell);
        } else if (theCell.find("input[type='radio']").length > 0) {
            fields.disableRadioField(theCell);
        } else if (theCell.find("span[class='ms-RadioText']>input[type='checkbox']").length > 0) {
            fields.disableMultiSelectField(theCell);
        } else {
            theControls = theCell.find("input,select,textarea,img");
            value = "<span class='readonly'>" + theControls.val() + "<span>";
            theControls.hide();
            theCell.prepend(value);
            theCell.find("div.ms-inputBox").hide();
        }
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
        selectedValues.each(() => {
            const selectedLabel = $(this).siblings("label");
            value += "<span class='readonly'>" + selectedLabel.text() + "<span><br/>";
        })
        theControls.hide();
        theCell.prepend(value);
    },

    disableRadioField: (theCell) => {
        //debugger
        const theControls = theCell.find("table");
        const selectedValue = theCell.find("input:checked").val();
        theControls.hide();
        theCell.prepend(selectedValue);
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

    getRow: (fieldDisplayName) => {
        let theRow = $("[Title='" + fieldDisplayName + "'], [Title='" + fieldDisplayName + " possible values']").closest("td.ms-formbody").parent();
        if (theRow.length == 0) {
            theRow = $(".ms-formlabel").filter(() => {
                return $(this).text().trim() == fieldDisplayName || $(this).text().trim() == fieldDisplayName + " *";
            }).parent();
        }
        return theRow;
    },

    getCell: (fieldDisplayName) => {
        let theCell = $("[Title='" + fieldDisplayName + "'], [Title='" + fieldDisplayName + " possible values']").closest("td.ms-formbody");
        if (theCell.length == 0) {
            theCell = $(".ms-formlabel").filter(() => {
                return $(this).text().trim() == fieldDisplayName || $(this).text().trim() == fieldDisplayName + " *";
            }).parent().find("td.ms-formbody");
        }
        return theCell;
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