# SharePoint-HideDisableFields
Lightweight JavaScript file to easily hide or disable fields on a SharePoint classic list form, supports SharePoint 2013, 2016 and Online.

## Usage ##
See https://davidlozzi.com/2014/01/14/sharepoint-2013-script-hide-or-disable-your-fields/ for usage and support.

Download the only listed JavaScript file above :D




#### Lozzi.Fields.disable(fieldname)

Simply disables the field, for all users. It hides all controls in the field and displays the value instead.

#### Lozzi.Fields.disableWithAllowance(fieldname, groups)

Disables the field, but enables it for the users in the groups specified. Also, Site Collection Administrators are included automatically, so they can always edit the field. You can send the groups in an array, like [“Group One”, “Group Two”].

#### Lozzi.Fields.hide(fieldname)

Simply hides the field, for all users.

#### Lozzi.Fields.hideWithAllowance(fieldname, groups)

Hides the field, but shows it for the users in the groups specified. Also, Site Collection Administrators are included automatically, so they can always edit the field. You can send the groups in an array, like [“Group One”, “Group Two”].

#### Some other important notes

Currently, this script does not work on list views, meaning a user could edit the data in datasheet/quick edit view.
This script should work just as well on SharePoint 2010 if you so desire.

### Support ###
[Post an issue here](https://github.com/DavidLozzi/SharePoint-HideDisableFields/issues)!
