import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFuelFormProps {
  hasTeamsContext: boolean;
  context: WebPartContext;
  companyNames: string[];
  suppliers: any[];
  targetListId: string;
}
