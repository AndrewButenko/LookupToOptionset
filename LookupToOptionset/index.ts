import { IDropdownOption } from "@fluentui/react";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { RecordSelector } from "./RecordSelector";

export class LookupToOptionset implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private container: HTMLDivElement;
	private notifyOutputChanged: () => void;
	private entityName: string;
	private availableOptions: IDropdownOption[];
	private currentValue?: ComponentFramework.LookupValue[];

	constructor() {
	}

	public init(context: ComponentFramework.Context<IInputs>,
		notifyOutputChanged: () => void,
		state: ComponentFramework.Dictionary,
		container: HTMLDivElement): void {
		this.container = container;
		this.notifyOutputChanged = notifyOutputChanged;

		this.entityName = context.parameters.lookup.getTargetEntityType();

		context.utils.getEntityMetadata(this.entityName).then(metadata => {
			const entityIdFieldName = metadata.PrimaryIdAttribute;
			const entityNameFieldName = metadata.PrimaryNameAttribute;

			const query = `?$select=${entityIdFieldName},${entityNameFieldName}`;

			context.webAPI.retrieveMultipleRecords(this.entityName, query).then(result => {
				this.availableOptions = result.entities.map(r => {
					return {
						key: r[entityIdFieldName],
						text: r[entityNameFieldName] ?? "Display Name is not available"
					};
				});

				this.renderControl(context);
			});
		});
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this.renderControl(context);
	}

	private renderControl(context: ComponentFramework.Context<IInputs>) {
		const recordId = context.parameters.lookup.raw != null && context.parameters.lookup.raw.length > 0 ?
			context.parameters.lookup.raw[0].id : undefined;

		const recordSelector = React.createElement(RecordSelector, {
			selectedRecordId: recordId,
			availableOptions: this.availableOptions,
			onChange: (selectedOption?: IDropdownOption) => {
				if (typeof selectedOption === "undefined") {
					this.currentValue = undefined;
				} else {
					this.currentValue = [{
						id: <string>selectedOption.key,
						name: selectedOption.text,
						entityType: this.entityName
					}];
				}

				this.notifyOutputChanged();
			}
		});

		ReactDom.render(recordSelector, this.container);
	}

	public getOutputs(): IOutputs {
		return {
			lookup: this.currentValue
		};
	}

	public destroy(): void {
		ReactDom.unmountComponentAtNode(this.container);
	}
}
