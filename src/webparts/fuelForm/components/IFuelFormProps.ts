import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFuelFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  spolkiNames: string[];
  dostawcy: any[];
}
