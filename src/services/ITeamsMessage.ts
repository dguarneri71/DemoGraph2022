export interface ITeamsMessage {
    body: IBody;
    createdDateTime: Date;
    deletedDateTime?: Date;
    user: string;
    messageType: string;
    channelIdentity: IChannelIdentity;
}

export interface IBody {
    content: string;
    contentType: string;
}

export interface IChannelIdentity {
    channelId: string;
    teamId: string;
}