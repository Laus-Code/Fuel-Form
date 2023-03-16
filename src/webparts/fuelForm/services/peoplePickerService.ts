import {IPersonaProps} from "office-ui-fabric-react";
import {graphfi, SPFx} from "@pnp/graph";
import "@pnp/graph/users"
import {WebPartContext} from "@microsoft/sp-webpart-base";

export class PeoplePickerService {
    public static async findPeople(filterText: string, context: WebPartContext): Promise<IPersonaProps[]> {
        const users: IPersonaProps[] = [];

        const graph = graphfi().using(SPFx(context));

        graph.users.search(filterText)().then((response) => {
            response.forEach((user) => {
                users.push({
                    id: user.id,
                    text: user.displayName,
                    secondaryText: user.mail
                })
            });
        }).catch(exception => console.log(exception));

        return users;
    }
}