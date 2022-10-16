import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClient } from "@microsoft/sp-http";
import { graph } from '@pnp/graph';
import "@pnp/graph/users";
import "@pnp/graph/teams";
import { IChannels, ITeams } from '@pnp/graph/teams';
import { ITeamsMessage } from './ITeamsMessage';
import { forEach } from 'lodash';

//usato pnp/pnpjs 2.14.0 non la V3.0
export default class SharePointDataService {
    private context: WebPartContext = null;

    public constructor(context: WebPartContext) {
        this.context = context;
        graph.setup({
            spfxContext: context
        });
    }

    public getMyTeams(): Promise<ITeams[]> {
        return graph.me.joinedTeams();
    }

    public getChannels(teamsID: string): Promise<IChannels[]> {
        return graph.teams.getById(teamsID).channels.get();
    }

    public async getMessages(teamsID: string, channelID: string): Promise<ITeamsMessage[]> {
        let messages: ITeamsMessage[] = [];
        let results = [];
        let client: MSGraphClient = await this.context.msGraphClientFactory.getClient();
        //let result: any = await client.api("https://graph.microsoft.com/beta/teams/" + teamsID + "/channels/" + channelID + "/messages").get();
        let result: any = await client.api(`teams/${teamsID}/channels/${channelID}/messages`).version("beta").get();
        console.log("SPDataService - getMessages() - result: ", result);
        if (result["value"]) {
            results = result["value"];
            results.forEach(element => {
                messages.push({
                    body: element["body"],
                    createdDateTime: new Date(element["createdDateTime"]),
                    deletedDateTime: element["deletedDateTime"] ? new Date(element["deletedDateTime"]) : null,
                    messageType: element["messageType"],
                    user: element["from"] ? element["from"]["user"]["displayName"] : "",
                    channelIdentity: element["channelIdentity"]
                });
            });
        }

        return new Promise<ITeamsMessage[]>(res => {
            res(messages);
        });
    }

    public async sendMessage(teamsID: string, channelID: string, msg: string): Promise<any> {
        let chatMessage = { body: { content: msg } };
        let client: MSGraphClient = await this.context.msGraphClientFactory.getClient();
        //let result = await client.api(`https://graph.microsoft.com/v1.0/teams/${teamsID}/channels/${channelID}/messages`).post(chatMessage);
        return await client.api(`teams/${teamsID}/channels/${channelID}/messages`).post(chatMessage);
    }
}