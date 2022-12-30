import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFuelFormProps {
  hasTeamsContext: boolean;
  title: string,
  maxFuelLimit: number,
  context: WebPartContext;
  companyNames: string[];
  suppliers: any[];
  targetListId: string;
}
