import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { IColumnReturnProperty, PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';
import * as strings from 'PropertyFieldsDemoWebPartStrings';
import PropertyFieldsDemo from './components/PropertyFieldsDemo';
import { IPropertyFieldsDemoProps } from './components/IPropertyFieldsDemoProps';

export interface IPropertyFieldsDemoWebPartProps {
    list: string;
    columnSingleTitle: string;
    columnMultipleID: string[];
    columnInternalName: string[];
}

export default class PropertyFieldsDemoWebPart extends BaseClientSideWebPart<IPropertyFieldsDemoWebPartProps> {

    public render(): void {
        const element: React.ReactElement<IPropertyFieldsDemoProps> = React.createElement(
            PropertyFieldsDemo,
            {
                list: this.properties.list,
                columnSingleTitle: this.properties.columnSingleTitle,
                columnTitle: this.properties.columnMultipleID,
                columnInternalName: this.properties.columnInternalName
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: 'List Selection',
                            groupFields: [
                                PropertyFieldListPicker('list', {
                                    label: 'Select a list',
                                    selectedList: this.properties.list,
                                    includeHidden: false,
                                    orderBy: PropertyFieldListPickerOrderBy.Title,
                                    disabled: false,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    context: this.context,
                                    onGetErrorMessage: null,
                                    deferredValidationTime: 0,
                                    key: 'listPickerFieldId'
                                }),


                            ]
                        },
                        {
                            groupName: 'Single Column Selection',
                            groupFields: [
                                PropertyFieldColumnPicker('columnSingleTitle', {
                                    label: 'Select a single column',
                                    context: this.context,
                                    selectedColumn: this.properties.columnSingleTitle,
                                    listId: this.properties.list,
                                    disabled: false,
                                    orderBy: PropertyFieldColumnPickerOrderBy.Title,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    onGetErrorMessage: null,
                                    deferredValidationTime: 0,
                                    key: 'columnSingleTitlePickerFieldId',
                                    displayHiddenColumns: false,
                                    columnReturnProperty: IColumnReturnProperty.Id
                                })
                            ]
                        },
                        {
                            groupName: 'MultiColumn Title Selection',
                            groupFields: [
                                PropertyFieldColumnPicker('columnMultipleID', {
                                    label: 'Select columns which will return title',
                                    context: this.context,
                                    selectedColumn: this.properties.columnMultipleID,
                                    listId: this.properties.list,
                                    disabled: false,
                                    orderBy: PropertyFieldColumnPickerOrderBy.Title,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    onGetErrorMessage: null,
                                    deferredValidationTime: 0,
                                    key: 'multiColumntitlePickerFieldId',
                                    displayHiddenColumns: false,
                                    columnReturnProperty: IColumnReturnProperty.Id,
                                    multiSelect: true
                                }),
                            ]
                        },
                        {
                            groupName: 'MultiColumn Internal Names selection',
                            groupFields: [
                                PropertyFieldColumnPicker('columnInternalName', {
                                    label: 'Select columns which will return internal names',
                                    context: this.context,
                                    selectedColumn: this.properties.columnInternalName,
                                    listId: this.properties.list,
                                    disabled: false,
                                    orderBy: PropertyFieldColumnPickerOrderBy.Title,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    onGetErrorMessage: null,
                                    deferredValidationTime: 0,
                                    key: 'multiColumninternalnamePickerFieldId',
                                    displayHiddenColumns: false,
                                    columnReturnProperty: IColumnReturnProperty['Internal Name'],
                                    multiSelect: true
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
