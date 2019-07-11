import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class AutoNumberControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _labelAutoNumber: HTMLLabelElement;
    private _container: HTMLDivElement;
    private _autoNumberDiv: HTMLDivElement;
    private _refreshDiv: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _imgRefresh: HTMLImageElement;
    private _refreshAutoNumber: EventListenerOrEventListenerObject;

	/**
	 * Empty constructor.
	 */
    constructor() {

    }

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
        this._context = context;
        this._container = container;

        this._autoNumberDiv = document.createElement("div");
        this._refreshDiv = document.createElement("div");
        this._labelAutoNumber = document.createElement("label");
        this._imgRefresh = document.createElement("img");

        this._container.appendChild(this._autoNumberDiv);
        this._container.appendChild(this._refreshDiv);
        this._autoNumberDiv.appendChild(this._labelAutoNumber);
        this._refreshDiv.appendChild(this._imgRefresh);

        this._container.setAttribute("class", "container");
        this._autoNumberDiv.setAttribute("class", "autonumberdiv");
        this._refreshDiv.setAttribute("class", "refreshdiv");
        this._imgRefresh.setAttribute("class", "refreshimage");
        
        this._refreshAutoNumber = this.generateAutoNumber.bind(this, true, context);
        this._imgRefresh.addEventListener("click", this._refreshAutoNumber);
    }


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        this.generateAutoNumber(false, context);
    }

    public generateAutoNumber(isRefresh: boolean, context: ComponentFramework.Context<IInputs>): void {
        // @ts-ignore
        let formType: string = Xrm.Page.ui.getFormType();
        // @ts-ignore
        let crmAutoNumberFieldName: string = context.parameters.fieldNameProperty.attributes.LogicalName;
        // @ts-ignore
        let entityName: string = Xrm.Page.data.entity.getEntityName();
        // @ts-ignore
        let recordId: string = Xrm.Page.data.entity.getId().replace("{", "").replace("}", "");
        let pluralEntityName: string = this.getEntityPluralName(entityName);
        let existingAutoNumber: string = this.getFieldValue(entityName, pluralEntityName, recordId, crmAutoNumberFieldName);

        if (isRefresh || (formType != "1" && existingAutoNumber == "")) {
            let autoNumber: string = this.processAutoNumber(entityName, pluralEntityName, recordId);
            if (autoNumber != existingAutoNumber) {
                this.saveRecord(pluralEntityName, recordId, crmAutoNumberFieldName, autoNumber);
            }
            this._labelAutoNumber.innerHTML = autoNumber;
        }
        else {
            this._labelAutoNumber.innerHTML = existingAutoNumber;
        }
    }

    public processAutoNumber(entityName: string, pluralEntityName: string, recordId: string): string {
        let resultList: string[] = [];
        let autoNumberFormat: string = this._context.parameters.formatProperty.raw;

        for (var i = 0; i < autoNumberFormat.length; i++) {
            let inputString: string = "";
            if (autoNumberFormat.charAt(i) == '[') {
                i++;
                while (autoNumberFormat.charAt(i) != ']') {
                    inputString += autoNumberFormat.charAt(i);
                    i++;
                }

                if (inputString.indexOf('(') > 0) {
                    let fieldName: string = inputString.split('(')[0].toString();
                    let refFieldEntityNameAndId = this.getFieldEnityNameAndRecordId(entityName, pluralEntityName, recordId, fieldName);
                    let refEntityName: string = refFieldEntityNameAndId.split('@')[0];
                    let refRecordId: string = refFieldEntityNameAndId.split('@')[1];
                    let refEntityPluralName: string = this.getEntityPluralName(refEntityName);
                    let refEntityFieldName: string = inputString.split('(')[1].toString().slice(0, -1);

                    resultList.push(this.getFieldValue(refEntityName, refEntityPluralName, refRecordId, refEntityFieldName));
                }
                else {
                    resultList.push(this.getFieldValue(entityName, pluralEntityName, recordId, inputString));
                }
            }
            else if (autoNumberFormat.charAt(i) == '{') {
                i++;
                while (autoNumberFormat.charAt(i) != '}') {
                    inputString += autoNumberFormat.charAt(i);
                    i++;
                }
                resultList.push(inputString);
            }
            else {
                while (autoNumberFormat.charAt(i) != '[' && autoNumberFormat.charAt(i) != '{') {
                    inputString += autoNumberFormat.charAt(i);
                    i++;
                }
                i--;
                resultList.push(inputString);
            }
        }
        return resultList.join('');
    }

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
    public getOutputs(): IOutputs {
        return {};
    }

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }

    public getFieldDataType(entityName: string, fieldName: string): string {
        let attributeType: string = "";
        var req = new XMLHttpRequest();
        // @ts-ignore
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/EntityDefinitions(LogicalName='" + entityName + "')/Attributes(LogicalName='" + fieldName + "')?$select=AttributeType", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                if (this.status == 200) {
                    var result = JSON.parse(this.response);
                    if (result != undefined && result["AttributeType"] != null) {
                        attributeType = result["AttributeType"];
                    }
                } else {
                    //Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
        return attributeType;
    }

    public getEntityPluralName(entityName: string): string {
        let entityPluralName: string = "";
        var req = new XMLHttpRequest();
        // @ts-ignore
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/EntityDefinitions(LogicalName='" + entityName + "')?$select=LogicalCollectionName", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState == 4) {
                req.onreadystatechange = null;
                if (this.status == 200) {
                    var result = JSON.parse(this.response);
                    if (result != undefined && result["LogicalCollectionName"] != null) {
                        entityPluralName = result["LogicalCollectionName"];
                    }
                } else {
                    //Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
        return entityPluralName;
    }

    public getFieldValue(entityName: string, entityPluralName: string, recordId: string, fieldName: string): string {
        let fieldDataType: string = this.getFieldDataType(entityName, fieldName);
        let fieldValue: string = "";
        if (fieldDataType != "") {
            let requestFieldName: string = fieldName;
            if (fieldDataType == "Lookup") {
                requestFieldName = "_" + fieldName + "_value";
            }
            var req = new XMLHttpRequest();
            // @ts-ignore
            req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/" + entityPluralName + "(" + recordId + ")?$select=" + requestFieldName, false);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
            req.onreadystatechange = function () {
                if (this.readyState == 4) {
                    req.onreadystatechange = null;
                    if (this.status == 200) {

                        let responseFieldName: string = requestFieldName;
                        if (fieldDataType == "Lookup" || fieldDataType == "Picklist" || fieldDataType == "Boolean" || fieldDataType == "Customer") {
                            responseFieldName += "@OData.Community.Display.V1.FormattedValue";
                        }
                        var result = JSON.parse(this.response);
                        if (result != undefined && result[responseFieldName] != null) {
                            fieldValue = result[responseFieldName];
                        }
                    } else {
                        //Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
        return fieldValue;
    }

    public getFieldEnityNameAndRecordId(entityName: string, entityPluralName: string, recordId: string, fieldName: string): string {
        let fieldDataType: string = this.getFieldDataType(entityName, fieldName);
        let fieldEntityNameAndRecordId: string = "";
        if (fieldDataType != "" && (fieldDataType == "Lookup" || fieldDataType == "Customer")) {
            let requestFieldName: string = "_" + fieldName + "_value";

            var req = new XMLHttpRequest();
            // @ts-ignore
            req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/" + entityPluralName + "(" + recordId + ")?$select=" + requestFieldName, false);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
            req.onreadystatechange = function () {
                if (this.readyState == 4) {
                    req.onreadystatechange = null;
                    if (this.status == 200) {

                        let responseFieldName: string = requestFieldName;
                        responseFieldName = requestFieldName + "@Microsoft.Dynamics.CRM.lookuplogicalname";
                        var result = JSON.parse(this.response);
                        if (result != undefined && result[responseFieldName] != null) {
                            fieldEntityNameAndRecordId = result[responseFieldName] + "@" + result[requestFieldName];
                        }
                    } else {
                        //Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
        return fieldEntityNameAndRecordId;
    }

    public saveRecord(entityName: string, recordId: string, fieldName: string, fieldValue: string): void {
        var entity = {};
        // @ts-ignore
        entity[fieldName] = fieldValue;

        var req = new XMLHttpRequest();
        // @ts-ignore
        req.open("PATCH", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/" + entityName + "(" + recordId + ")", true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 204) {
                    //Success - No Return Data - Do Something
                } else {
                    //Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send(JSON.stringify(entity));
    }
}