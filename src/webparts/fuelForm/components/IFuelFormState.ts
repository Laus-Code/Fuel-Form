import { IPersonaProps } from 'office-ui-fabric-react';

export interface IFuelFormState {
  formForUser?: boolean;
  personOnList?: boolean;
  driver?: IPersonaProps[];
  name?: string;
  surname?: string;
  email?: string;
  supervisor?: IPersonaProps[];
  company?: string;
  registrationNumber?: string;
  supplier?: string;
  mask?: string;
  cardNumber?: string;
  distance?: number;
  limitChange?: number;
  route?: string;
  startDate?: Date;
  endDate?: Date;
  justification?: string;
}