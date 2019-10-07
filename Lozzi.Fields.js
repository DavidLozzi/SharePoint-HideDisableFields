"use strict"

/*
Created by David Lozzi, @davidlozzi, www.davidlozzi.com, 1/10/2014
Last Updated 7/5/2019
Requires jQuery, you can include: <script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
*/

var fields = {
  spContext: null,
  currentUser: null,
  peopleLoopCount: 0,
  peoplePickerBorder: "",
  metadataLoop: 0,

  getFieldType: function (theCell) {
    var fieldType = '';
    theCell.contents().each(function (i, obj) {
      if (obj.nodeType === 8) {
        var ftRegx = new RegExp('(FieldType=")([a-zA-Z]*)"', 'gmi');
        var vals = ftRegx.exec(obj.data);
        if (vals && vals.length > 2) {
          fieldType = vals[2];
        }
      }
    });
    return fieldType;
  },

  getRow: function (fieldDisplayName) {
    var theCell = fields.getCell(fieldDisplayName);
    var theRow = theCell.parent();
    return theRow;
  },

  getCell: function (fieldDisplayName) {
    var theCell = $(".ms-formbody").filter(function (i, o) {
      var rx = 'FieldName=\"' + fieldDisplayName + '\"';
      var html = o.innerHTML.match(new RegExp(rx));
      return html !== null; // && html.length > 0;
    });
    return theCell;
  },

  writeValue: function (theCell, value) {
    var spanny = "<span class='readonly'>" + value + "</span>";
    theCell.prepend(spanny);
  },

  disable: function (fieldDisplayName) {
    var theCell = fields.getCell(fieldDisplayName);
    var fieldType = fields.getFieldType(theCell);

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
      case 'SPFieldUserMulti':
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

  defaultType: function (theCell) {
    var theControls = theCell.find("input,select,textarea,img");
    fields.writeValue(theCell, theControls.val());

    theControls.hide();
    theCell.find("div.ms-inputBox").hide();
  },

  disableOtherOption: function (theCell) {
    var theControls = theCell.find("input,select,textarea,img,label,table");
    var fillInValue = theCell.find("input[id$='FillInChoice'], input[id$='FillInText']").val();
    fields.writeValue(theCell, fillInValue);
    theControls.hide();
  },

  disablePeoplePicker: function (theCell) {
    if ($(".sp-peoplepicker-topLevel", theCell).css("border") != "none") {
      //console.log("disablePeoplePicker")
      $(".sp-peoplepicker-editorInput", theCell).attr("disabled", "true")
      fields.peoplePickerBorder = $(".sp-peoplepicker-topLevel", theCell).css("border");
      $(".sp-peoplepicker-topLevel", theCell).css("border", "none");
      $(".sp-peoplepicker-delImage", theCell).hide();

      fields.peopleLoopCount++;
      //hate this loop, not sure how else to handle waiting for the objs 
      if (fields.peopleLoopCount < 3) {
        setTimeout(function () { fields.disablePeoplePicker(theCell) }, 350);
      } else {
        fields.peopleLoopCount = 0;
      }
    }
  },

  disableMetadata: function (theCell) {
    if ($(".valid-text", theCell).length > 0 || fields.metadataLoop == 5) {
      //console.log("found " + $(".valid-text",theCell).length + " with text " + $(".valid-text",theCell).text());
      var metaValue = $(".valid-text", theCell).text();
      var theControls = theCell.find("input,select,textarea,img");
      fields.writeValue(theCell, metaValue);
      theControls.hide();

      theCell.find("div.ms-inputBox").hide();
      fields.metadataLoop = 0;
    } else {
      fields.metadataLoop++;
      //yup, hate looping to wait for the DOM to load. Must. Find. Fix.
      if (fields.metadataLoop <= 5) {
        setTimeout(function () { fields.disableMetadata(theCell) }, 350);
      }
    }
  },

  disableLookupField: function (theCell) {
    var theControls = theCell.find("select");
    var selectedValue = theControls.find("option:selected");
    fields.writeValue(theCell, selectedValue.text());
    theControls.hide();
  },

  disableMultiSelectField: function (theCell) {
    var theControls = theCell.find("table");
    var selectedValues = theControls.find("input:checked");
    var value = "";
    selectedValues.each(function (i, o) {
      var selectedLabel = $(o).siblings("label");
      if (selectedLabel.text() === "Specify your own value:") {
        value += theCell.find("input[id$=FillInText]").val() + "<br/>";
      } else {
        value += selectedLabel.text() + "<br/>";
      }
    })
    fields.writeValue(theCell, value);
    theControls.hide();
  },

  disableYesNo: function (theCell) {
    var theControls = theCell.find("input");
    var isChecked = theControls.is(":checked");
    var value = isChecked ? "Yes" : "No";
    fields.writeValue(theCell, value);
    theControls.hide();
  },

  disableRadioField: function (theCell) {
    var theControls = theCell.find("table");
    var selectedValue = theCell.find("input:checked").val();
    if (selectedValue === "DropDownButton") {
      selectedValue = theCell.find("option:selected").val();
    }
    var value = selectedValue + "<br/>";
    fields.writeValue(theCell, value);
    theControls.hide();
  },

  enable: function (fieldDisplayName) {
    var theCell = fields.getCell(fieldDisplayName);
    var fieldType = fields.getFieldType(theCell);

    switch (fieldType) {
      case 'SPFieldTaxonomyFieldType':
        fields.enableMetadata(theCell);
        break;
      case 'SPFieldUser':
      case 'SPFieldUserMulti':
        fields.enablePeoplePicker(theCell);
        break;
      default:
        var theControls = theCell.find("input,select,textarea,img");
        var value = theCell.find("span.readonly")
        theControls.show();
        value.remove();
        break;
    }
  },

  enablePeoplePicker: function (theCell) {
    if ($(".sp-peoplepicker-topLevel", theCell).css("border") != fields.peoplePickerBorder) {
      //console.log("enablePeoplePicker")
      $(".sp-peoplepicker-editorInput", theCell).removeAttr("disabled")
      $(".sp-peoplepicker-topLevel", theCell).css("border", fields.peoplePickerBorder);
      $(".sp-peoplepicker-delImage", theCell).show();
      $("span.readonly", theCell).remove();
      fields.peopleLoopCount++;
      if (fields.peopleLoopCount < 7) {
        setTimeout(function () { fields.enablePeoplePicker(theCell) }, 350);
      } else {
        fields.peopleLoopCount = 0;
      }
    }
  },

  enableMetadata: function (theCell) {
    console.log("enableMetadata");
    if (fields.metadataLoop == 0) {
      //console.log("found " + $(".valid-text",theCell).length + " with text " + $(".valid-text",theCell).text());
      var theControls = theCell.find("input,select,textarea,img");
      theControls.show();
      $("div.ms-inputBox", theCell).show();
      $("span.readonly", theCell).remove();
    } else {
      setTimeout(function () { fields.enableMetadata(theCell) }, 350);
    }
  },

  hide: function (fieldDisplayName) {
    var theRow = fields.getRow(fieldDisplayName);
    theRow.hide();
  },

  show: function (fieldDisplayName) {
    var theRow = fields.getRow(fieldDisplayName);
    theRow.show();
  },

  //group names in an array of strings, i.e. ["Group One","Group Two"]
  disableWithAllowance: function (fieldName, groups) {
    fields.disable(fieldName);

    fields.spContext = new SP.ClientContext.get_current();
    fields.currentUser = fields.spContext.get_web().get_currentUser();

    fields.spContext.load(fields.currentUser);
    fields.spContext.load(fields.currentUser.get_groups());
    fields.spContext.executeQueryAsync(function () {
      fields.getGroupsAndEnable(fieldName, groups)
    }, fields.asyncFailed);
  },

  getGroupsAndEnable: function (fieldName, groups) {
    var allowedToEdit = false;
    if (fields.currentUser.get_isSiteAdmin()) {
      allowedToEdit = true;
    } else {
      var groupEnum = fields.currentUser.get_groups().getEnumerator();
      while (groupEnum.moveNext()) {
        var group = groupEnum.get_current();
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
  hideWithAllowance: function (fieldName, groups) {
    fields.hide(fieldName);

    fields.spContext = new SP.ClientContext.get_current();
    fields.currentUser = fields.spContext.get_web().get_currentUser();

    fields.spContext.load(fields.currentUser);
    fields.spContext.load(fields.currentUser.get_groups());
    fields.spContext.executeQueryAsync(function () {
      fields.getGroupsAndShow(fieldName, groups);
    }, fields.asyncFailed);
  },

  getGroupsAndShow: function (fieldName, groups) {
    var allowedToEdit = false;
    if (fields.currentUser.get_isSiteAdmin()) {
      allowedToEdit = true;
    } else {
      var groupEnum = fields.currentUser.get_groups().getEnumerator();
      while (groupEnum.moveNext()) {
        var group = groupEnum.get_current();
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

  asyncFailed: function (sender, args) {
    alert('request failed ' + args.get_message() + '\n' + args.get_stackTrace());
  },

  setDefaultValue: function (fieldDisplayName, stringValue) {
    //TODO flush out the many field type options
    var theCell = fields.getCell(fieldDisplayName);
    var selectObj = theCell.find("SELECT");
    selectObj.val(stringValue);
  },

  formatRow: function (fieldDisplayName, css) {
    var theRow = fields.getRow(fieldDisplayName);
    theRow.css(css);
  },

  formatLabel: function (fieldDisplayName, css) {
    var theRow = fields.getRow(fieldDisplayName);
    theRow.find(".ms-formlabel").css(css);
  },

  formatValueCell: function (fieldDisplayName, css) {
    var theCell = fields.getCell(fieldDisplayName);
    theCell.css(css);
  }
};

var Lozzi = window.Lozzi || {};
Lozzi.Fields = fields;