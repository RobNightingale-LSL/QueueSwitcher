import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class QueueSwitcher implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    //private _createQueueSwitchButton: HTMLButtonElement;
    private static _queueItemEntityName: string = "queueitem";
    private _currentEntityName: string;
    private _currentEntityId: string;
    private _userId: string;
    
    private _resultDivContainer: HTMLDivElement;

    private _controlViewRendered: Boolean;
    private notifyOutputChanged: () => void;

	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
    public init(context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement) {
        this._context = context;
        this._controlViewRendered = false;
        this._container = document.createElement("div");
        this._container.classList.add("QueueSwitcher_Container");
        container.appendChild(this._container);
        //TODO get this supoorted
        //@ts-ignore
        this._currentEntityId = this._context.mode.contextInfo.entityId;
        //@ts-ignore
        this._currentEntityName = this._context.mode.contextInfo.entityTypeName;
        this._userId = this._context.userSettings.userId;
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
        if (!this._controlViewRendered) {
            this._controlViewRendered = true;
            this.renderSwitcherDiv();
            this.renderResultsDiv();
        }
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		//TODO: Ask Andrii if any is needed here
    }

    private renderResultsDiv() {
        // Render header label for result container
        /*let resultDivHeader: HTMLDivElement = this.createHTMLDivElement(
            "result_container",
            true,
            "Result of last action"
        );*/
        // Div elements to populate with the result text
        this._resultDivContainer = this.createHTMLDivElement(
            "result_container",
            false,
            undefined
        );

        //this._container.appendChild(resultDivHeader);
        this._container.appendChild(this._resultDivContainer);

        // Init the result container with a notification the control was loaded
        this.updateResultContainerText("Web API sample custom control loaded");
    }

    private renderSwitcherDiv() {
        let thisRef = this;
        debugger; //TODO: remove
        
        //TODO get top 5 queues from queueitems where worked by = current        
        thisRef._context.webAPI.retrieveMultipleRecords(QueueSwitcher._queueItemEntityName, "?$filter=_modifiedby_value eq " + thisRef._userId + "&$top=3&$orderby=modifiedon desc&$select=_queueid_value&$expand=queueid($select=queueid,name)").then(
            function (response: any) {
                debugger;//TODO: remove

                //TODO: loop the records for top 5 distinct values
                /*var grouped = response.entities.reduce((g: any, person: QueueItem) => {
                    g[person._queueid_value] = g[person._queueid_value] || []; //Check the value exists, if not assign a new array
                    g[person._queueid_value].push(person); //Push the new value to the array

                    return g; //Very important! you need to return the value of g or it will become undefined on the next pass
                }, {});*/

                response.entities.forEach((qi: any) => {
                    let createQueueSwitchButton = thisRef.createHTMLButtonElement(
                        "Move to " + qi.queueid["name"] + " Queue",
                        "qsButton-" + qi.queueid["queueid"],
                        qi.queueid["queueid"],
                        thisRef.createButtonOnClickHandler.bind(thisRef)
                    );
                    thisRef._container.appendChild(createQueueSwitchButton);
                });
            },
            function (errorResponse: any) {
                debugger;//TODO: remove
                // Error handling code here - record failed to be created
                thisRef.updateResultContainerTextWithErrorResponse(errorResponse);
            }
        );
        
        this.grabCurrentQueue();
    }

    /**
   * Helper method to create HTML Button used for CreateRecord Web API Example
   * @param buttonLabel : Label for button
   * @param buttonId : ID for button
   * @param buttonValue : value of button (attribute of button)
   * @param onClickHandler : onClick event handler to invoke for the button
   */
    private createHTMLButtonElement(
        buttonLabel: string,
        buttonId: string,
        buttonValue: string | null,
        onClickHandler: (event?: any) => void
    ): HTMLButtonElement {
        let button: HTMLButtonElement = document.createElement("button");
        button.innerHTML = buttonLabel;
        if (buttonValue) {
            button.setAttribute("buttonvalue", buttonValue);
        }
        button.id = buttonId;
        button.classList.add("QueueSwitcher_ButtonClass");
        button.addEventListener("click", onClickHandler);
        return button;
    }

    private grabCurrentQueue(): void {
        // store reference to 'this' so it can be used in the callback method
        var thisRef = this;
               
        //check for active
        this._context.webAPI.retrieveMultipleRecords(QueueSwitcher._queueItemEntityName, "?$select=queueitemid,_queueid_value&$filter=_objectid_value eq " + thisRef._currentEntityId + " and statecode eq 0&$expand=queueid($select=queueid,name)").then(
            function (response: any) {
                // Callback method for successful creation of new record
                debugger;//TODO: remove
                let resultHtml: string = "";
                if (Array.isArray(response.entities) && response.entities.length > 0) {
                    resultHtml = "Current Queue: " + response.entities[0].queueid.name;
                }
                else {
                    let resultHtml: string = "Not attached to a queue";
                }
                thisRef.updateResultContainerText(resultHtml);
            },
            function (errorResponse: any) {
                //debugger;
                // Error handling code here - record failed to be created
                thisRef.updateResultContainerTextWithErrorResponse(errorResponse);
            }
        );

    }

    /**
   * Event Handler for onClick of button
   * @param event : click event
   */
    private createButtonOnClickHandler(event: Event): void {
        var thisRef = this;
        debugger; //TODO: remove

        let queueGuid: string = (<HTMLInputElement>event.target!).attributes.getNamedItem("buttonvalue")!.value;      
        
        this.AddToQueue("AddToQueue", queueGuid, thisRef._currentEntityId, thisRef._currentEntityName);        
    }

    public AddToQueue(actionUniqueName: string, queueId: string, objectId: string, objectType: string): void {
        let thisRef = this;

        var request: any = new Object();

        var entity: any;
        entity = new Object();
        entity.id = queueId;
        entity.entityType = "queue";
        request.entity = entity;

        var target: any;
        target = new Object();
        target.incidentid = objectId;
        target["@odata.type"] = "Microsoft.Dynamics.CRM." + objectType;
        request.Target = target;

        request.getMetadata = function () {
            return {
                boundParameter: "entity",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.queue",
                        "structuralProperty": 5
                    },
                    "Target": {
                        "typeName": "mscrm.crmbaseentity",
                        "structuralProperty": 5
                    }
                },
                operationType: 0,
                operationName: actionUniqueName
            }
        };

        debugger; //TODO: remove
        if (request) {
            //@ts-ignore
            Xrm.WebApi.online.execute(request).then(
                function (result: any) {
                    debugger;//TODO: remove
                    thisRef.grabCurrentQueue();
                },
                function (error: any) {
                    debugger;//TODO: remove
                    thisRef.updateResultContainerTextWithErrorResponse(error.innerror.message);
                }
            );
        }
    }


    /**
   * Helper method to inject HTML into result container div
   * @param message : HTML to inject into result container
   */
    private updateResultContainerText(message: string): void {
        if (this._resultDivContainer) {
            this._resultDivContainer.innerHTML = message;
        }
    }
    /**
     * Helper method to inject error string into result container div after failed Web API call
     * @param errorResponse : error object from rejected promise
     */
    private updateResultContainerTextWithErrorResponse(errorResponse: any): void {
        if (this._resultDivContainer) {
            // Retrieve the error message from the errorResponse and inject into the result div
            let errorHTML: string = "Error with this control:";
            errorHTML += "<br />";
            errorHTML += errorResponse.message;
            this._resultDivContainer.innerHTML = errorHTML;
        }
    }
    /**
   * Helper method to create HTML Div Element
   * @param elementClassName : Class name of div element
   * @param isHeader : True if 'header' div - adds extra class and post-fix to ID for header elements
   * @param innerText : innerText of Div Element
   */
    private createHTMLDivElement(
        elementClassName: string,
        isHeader: boolean,
        innerText?: string): HTMLDivElement {
        let div: HTMLDivElement = document.createElement("div");
        if (isHeader) {
            div.classList.add("QueueSwitcher_Header");
            elementClassName += "_header";
        }
        if (innerText) {
            div.innerText = innerText.toUpperCase();
        }
        div.classList.add(elementClassName);
        return div;
    }
}