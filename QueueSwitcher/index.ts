import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { QueueSwitcherControl, IQueueSwitcherControlProps } from './QueueSwitcherControl';

export class QueueSwitcher implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private container: HTMLDivElement;
    private currentQueueId: string | null = null;
    private currentRecordEntityType: string;
    private currentRecordPKName: string;
    private currentRecordId: string | null = null;
    private webApi: ComponentFramework.WebApi;
    private navigation: ComponentFramework.Navigation;

    constructor() {
    }

    public init(context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement) {
        this.container = container;
        this.webApi = context.webAPI;
        this.navigation = context.navigation;
        this.currentRecordEntityType = (<any>context.mode).contextInfo.entityTypeName;
        
        let allPromises = [];
        //@ts-ignore
        allPromises.push(Xrm.Utility.getEntityMetadata(this.currentRecordEntityType));

        allPromises.push(context.webAPI.retrieveMultipleRecords("queue","?$select=name,queueid&$filter=queueviewtype eq 0"));

        if((<any>context.mode).contextInfo.entityId != null) {
            this.currentRecordId = (<any>context.mode).contextInfo.entityId;

            allPromises.push(context.webAPI.retrieveMultipleRecords("queueitem", "?$select=_queueid_value,queueitemid&$filter=_objectid_value eq " + this.currentRecordId));
        }

        Promise.all(allPromises).then((results) => {
            this.currentRecordPKName = results[0].PrimaryIdAttribute;

            let queueSwitcherProps: IQueueSwitcherControlProps = {
                Queues: results[1].entities.map((queue: { queueid: any; name: any; }) => {
                    return {
                        key: queue.queueid,
                        text: queue.name,
                        iconProps: {
                            iconName: "Assign"
                        }
                    };
                }),
                OnQueueChanged: this.moveToQueue,
                SelectedQueue: results.length === 2 || results[2].entities.length === 0 ? null : results[2].entities[0]._queueid_value
            };

            ReactDOM.render(React.createElement(QueueSwitcherControl, queueSwitcherProps), this.container);
        }, (e) => {
            console.log(e);
        });
    }

    private moveToQueue = (queueId: string) => {
        this.currentQueueId = queueId;

        if (this.currentRecordId == null) {
            return;
        }

        let target: any = {
            "@odata.type": "Microsoft.Dynamics.CRM." + this.currentRecordEntityType
        };
        target[this.currentRecordPKName] = this.currentRecordId;

        const addToQueueRequest = {
            entity: {
                id: queueId,
                entityType: "queue"
            },
            Target: target,
            getMetadata: function() {
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
                    operationName: "AddToQueue"
                };
            }
        };
        
        (<any>this.webApi).execute(addToQueueRequest).then(
            () => {},
            (error: any) => {
                this.navigation.openErrorDialog(error);
            }
        );    
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        if (context.updatedProperties.includes("entityId")) {
            this.currentRecordId = (<any>context.mode).contextInfo.entityId;
            
            this.currentQueueId && this.moveToQueue(this.currentQueueId);
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.container);
    }

    /*private renderSwitcherDiv() {
        let thisRef = this;

        thisRef._context.webAPI.retrieveMultipleRecords(QueueSwitcher._queueItemEntityName, "?$filter=_modifiedby_value eq " + thisRef._userId + "&$top=50&$orderby=modifiedon desc&$select=_queueid_value&$expand=queueid($select=queueid,name)").then(
            function (response: any) {

                let counter: QueueCount[] = [];

                response.entities.forEach((qi: any) => {
                    let found = counter.find(i => i.queueId === qi.queueid["queueid"])
                    if (found !== undefined) {
                        found.count++;
                    } else {
                        //create new object and add to array
                        let queue = new QueueCount(qi.queueid["name"], qi.queueid["queueid"], 1)
                        counter.push(queue);
                    }
                });

                counter.sort((a, b) => { return b.count - a.count });;

                let topX = counter.slice(0, thisRef._buttonCount);

                topX.forEach((qi: QueueCount) => {

                    let createQueueSwitchButton = thisRef.createHTMLButtonElement(
                        "Move to " + qi.queueName + " Queue",
                        "qsButton-" + qi.queueId,
                        qi.queueId,
                        thisRef.createButtonOnClickHandler.bind(thisRef)
                    );
                    thisRef._container.appendChild(createQueueSwitchButton);

                });
            },
            function (errorResponse: any) {
                // Error handling code here - record failed to be created
                thisRef.updateResultContainerTextWithErrorResponse(errorResponse);
            }
        );

        this.grabCurrentQueue();
    }*/

    /**
   * Helper method to create HTML Button used for CreateRecord Web API Example
   * @param buttonLabel : Label for button
   * @param buttonId : ID for button
   * @param buttonValue : value of button (attribute of button)
   * @param onClickHandler : onClick event handler to invoke for the button
   */
    /*private createHTMLButtonElement(
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
                thisRef.updateResultContainerTextWithErrorResponse(errorResponse);
            }
        );

    }*/

    /**
   * Event Handler for onClick of button
   * @param event : click event
   */
    /*private createButtonOnClickHandler(event: Event): void {
        var thisRef = this;

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

        if (request) {
            //@ts-ignore
            Xrm.WebApi.online.execute(request).then(
                function (result: any) {
                    thisRef.grabCurrentQueue();
                },
                function (error: any) {
                    thisRef.updateResultContainerTextWithErrorResponse(error.innerror.message);
                }
            );
        }
    }*/


    /**
   * Helper method to inject HTML into result container div
   * @param message : HTML to inject into result container
   */
    /*private updateResultContainerText(message: string): void {
        if (this._resultDivContainer) {
            this._resultDivContainer.innerHTML = message;
        }
    }*/
    /**
     * Helper method to inject error string into result container div after failed Web API call
     * @param errorResponse : error object from rejected promise
     */
    /*private updateResultContainerTextWithErrorResponse(errorResponse: any): void {
        if (this._resultDivContainer) {
            // Retrieve the error message from the errorResponse and inject into the result div
            let errorHTML: string = "Error with this control:";
            errorHTML += "<br />";
            errorHTML += errorResponse.message;
            this._resultDivContainer.innerHTML = errorHTML;
        }
    }*/
    /**
   * Helper method to create HTML Div Element
   * @param elementClassName : Class name of div element
   * @param isHeader : True if 'header' div - adds extra class and post-fix to ID for header elements
   * @param innerText : innerText of Div Element
   */
    /*private createHTMLDivElement(
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
    }*/
}