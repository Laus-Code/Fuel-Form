import {IPersonaProps} from "office-ui-fabric-react";

export interface IFuelFormState {
    formForUser?: boolean;
    personOnList?: boolean;
    success?: boolean;
    formSent?: boolean;
    showErrorBar?: boolean;
    showErrorMessages?: boolean;
    driver?: IPersonaProps[];
    name?: string;
    surname?: string;
    email?: string;
    supervisor?: IPersonaProps[];
    company?: string;
    registrationNumber?: string;
    supplier?: number;
    mask?: string;
    cardNumber?: string;
    distance?: number;
    limitChange?: number;
    route?: string;
    startDate?: Date;
    endDate?: Date;
    justification?: string;
}
