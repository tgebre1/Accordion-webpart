import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

//import jquery and jquerry UI
import * as jQuery from "jquery";
import "jqueryui";
import * as bootstrap from "bootstrap";

//Load some external CSS files by using the module loader. Add the following import:
import { SPComponentLoader } from "@microsoft/sp-loader";
//We need to import all the field types we want to use for Web Part pane properties
import {
IPropertyPaneConfiguration,
PropertyPaneTextField, //import single line and multiline text field
PropertyPaneButton, //buttons
PropertyPaneCheckbox, //check boxes
PropertyPaneDropdown, //dropdown menu
PropertyPaneToggle, //toggle button
PropertyPaneLabel,
PropertyPaneSlider,
IPropertyPaneDropdownOption,
PropertyPaneLink,
PropertyPaneHorizontalRule,
PropertyPaneButtonType

//* tip, click ctrl +space to get the intelisense /suggestions
} from "@microsoft/sp-property-pane";

// we will use this escape function from lodash to escape html special characters
import { escape } from "@microsoft/sp-lodash-subset";

//importing css
import styles from "./JQWebPartWebPart.module.scss";
import * as strings from "JQWebPartWebPartStrings";

//importing the mock array of items
import MockHttpClient from "./MockHttpClient";

//helper class SPHttpClient is provided by SharePoint to execute REST API Requests.
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

//SharePoint Framework aids this capability by helping you understand which
//environment your web part is running from by using the EnvironmentType module.
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

//creating a interface that will be used for assigning properties and type checking
export interface IjQWebPartProps {
webpartTitle: string;
siteURL: string;
sourceList: string;
headerColumnName: string;
contentColumnName: string;
noOfItems: number;
}

// list models to start working with SharePoint list data
//The ISPList interface holds the SharePoint list information that we are connecting to.
export interface ISPLists {
value: ISPList[];
}

export interface ISPList {
Header: string;
Description: string;
ID?: number;
}

//jQWebPart extends from class BaseClientSideWebPart with interface IjQWebPartProps using <>
export default class jQWebPart extends BaseClientSideWebPart<
IjQWebPartProps
> {
public constructor() {
super();

SPComponentLoader.loadCss(
"//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css"
);
SPComponentLoader.loadCss(
"//stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css"
);
SPComponentLoader.loadCss(
"//use.fontawesome.com/releases/v5.8.2/css/all.css"
);
}
//private method that mocks the list retrieval
private _getMockListData(): Promise<ISPLists> {
return MockHttpClient.get().then((data: ISPList[]) => {
var listData: ISPLists = { value: data };
return listData;
}) as Promise<ISPLists>;
}
//private method to pull data from accordian list in sharepoint online
private _getListData(): Promise<ISPLists> {
return this.context.spHttpClient
.get(
this.context.pageContext.web.absoluteUrl +
`/_api/web/lists/GetByTitle('${
this.properties.sourceList
}')/items?top=${this.properties.noOfItems}&$select=ID,${
this.properties.headerColumnName
},${this.properties.contentColumnName}`,
SPHttpClient.configurations.v1
)
.then((response: SPHttpClientResponse) => {
return response.json();
});
}
//private method to render list information received from REST API
private _renderList(items: ISPList[]): void {
let html: string = ``;

items.forEach((item: ISPList) => {
var url = `${this.context.pageContext.web.absoluteUrl}/lists/${
this.properties.sourceList
}/EditForm.aspx?ID=${item.ID}`;
html += `
<h3>${item.Header}</h3>
<div>
<p>${
item.Description
}<a class="float-right" target=\"_self"\ href=\"${url}"><i class="fas fa-edit"></i></a>
</p>
</div>`;
});
const listContainer: Element = this.domElement.querySelector(
"#spListContainer"
);
listContainer.innerHTML = html;

//add options
const accordionOptions: JQueryUI.AccordionOptions = {
animate: true,
collapsible: false,
icons: {
header: "ui-icon-circle-arrow-e",
activeHeader: "ui-icon-circle-arrow-s"
}
};
//initialize the accordian, class is .accordian and pass above options
jQuery(".accordion", this.domElement).accordion(accordionOptions);
}

//private method to call the respective methods to retrieve list data:
private _renderListAsync(): void {
// Local environment
if (Environment.type === EnvironmentType.Local) {
this._getMockListData().then(response => {
this._renderList(response.value);
});
} else if (
Environment.type == EnvironmentType.SharePoint ||
Environment.type == EnvironmentType.ClassicSharePoint
) {
this._getListData().then(response => {
this._renderList(response.value);
});
}
}

public render(): void {
this.domElement.innerHTML = `
<h3 class="">${escape(
this.context.pageContext.web.title
)}</h3>
<h4 class="p-3 mb-2 bg-dark text-white rounded">${
this.properties.webpartTitle
}</h4>
<div id="spListContainer" class="accordion"/>
</div>`;

this._renderListAsync();
}
//properties defined in IjQWebPartProps can be accessed using $(this.properties.nameOfProperty} )

//Note Above escape function imported from lodash library
// Notice that we are performing an HTML escape on the property's value to ensure a valid string
//escape Converts the characters "&", "<", ">", '"', and "'" in string to their corresponding HTML entities.

protected get dataVersion(): Version {
return Version.parse("1.0");
}

//method for validations
private validateFields(value: string): string {
if (value === null || value.trim().length === 0) {
return "Provide a value";
}

if (value.length > 200) {
return "Value should not be longer than 200 characters";
}

return "";
}
private cobWPPropButtonClick() {
alert("Property pane horozontal rule");
}

//In this method we add new properties to pane and mamp them to their typed objects
protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
return {
//we can add them to multiple pages
pages: [
{
header: {
description: "SharePoint News WebPart"
},
//properties can be defined into groups,
groups: [
{
groupName: "News Settings",
groupFields: [
PropertyPaneTextField("webpartTitle", {
label: "Web Part Title",
onGetErrorMessage: this.validateFields

//where is this coming from?

//ctrl + space gives all possible ptions you can use here
}),

PropertyPaneTextField("siteURL", {
label: "Site URL",
onGetErrorMessage: this.validateFields
}),
PropertyPaneTextField("sourceList", {
label: "Source List Name",
onGetErrorMessage: this.validateFields
}),

PropertyPaneTextField("headerColumnName", {
label: "Field Name for accordion header",
onGetErrorMessage: this.validateFields
}),
PropertyPaneTextField("contentColumnName", {
label: "Field Name for accordion body",
onGetErrorMessage: this.validateFields
}),
// ^ note both multine line and single line use same field type

PropertyPaneSlider("noOfItems", {
label: "Max no of items to display",
min: 1,
max: 100,
step: 2,
showValue: true
//ctrl + space gives all possible options you can use here
}),
PropertyPaneHorizontalRule(),
PropertyPaneButton("", {
text: "View Details",
buttonType: PropertyPaneButtonType.Normal,
onClick: this.cobWPPropButtonClick
})
]
},
{
groupName: "Contact Us",
groupFields: [
PropertyPaneLink("", {
href: "https://www.google.com",
text: "Connect with us",
target: "_blank",
popupWindowProps: {
height: 500,
width: 500,
positionWindowPosition: 2,
title: "COB blog"
}
})
]
}
]
}
]
};
}
}