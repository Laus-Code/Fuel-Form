import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "FuelFormWebPartStrings";
import FuelForm from "./components/FuelForm";
import { IFuelFormProps } from "./components/IFuelFormProps";

import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";

import { getSP } from "./components/pnpjs-config";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { SPFI } from "@pnp/sp";

export interface IFuelFormWebPartProps {
  description: string;
  companyListId: string;
  supplierListId: string;
  targetListId: string;
}

export default class FuelFormWebPart extends BaseClientSideWebPart<IFuelFormWebPartProps> {
  private companyNames: string[];
  private suppliers: any[];
  private sp: SPFI;

  public async render(): Promise<void> {
    await this.getListsData();

    const element: React.ReactElement<IFuelFormProps> = React.createElement(
      FuelForm,
      {
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        context: this.context,
        companyNames: this.companyNames,
        suppliers: this.suppliers,
        targetListId: this.properties.targetListId,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this.sp = getSP(this.context);
    return super.onInit();
  }

  private async getListsData(): Promise<void> {
    if (this.properties.companyListId) {
      this.companyNames = (
        await this.sp.web.lists.getById(this.properties.companyListId).items()
      ).map((s) => s.Title);
    }

    if (this.properties.supplierListId) {
      this.suppliers = await this.sp.web.lists
        .getById(this.properties.supplierListId)
        .items();
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: "Listy",
              groupFields: [
                PropertyFieldListPicker("supplierListId", {
                  label: "Wybierz listę z dostawcami",
                  selectedList: this.properties.supplierListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "supplierListPickerFieldId",
                }),
                PropertyFieldListPicker("companyListId", {
                  label: "Wybierz listę ze spółkami",
                  selectedList: this.properties.companyListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "companyListPickerFieldId",
                }),
                PropertyFieldListPicker("targetListId", {
                  label: "Wybierz listę docelową",
                  selectedList: this.properties.targetListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "targetlLstPickerFieldId",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
