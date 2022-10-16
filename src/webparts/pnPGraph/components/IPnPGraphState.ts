import { IChannels, ITeams } from "@pnp/graph/teams";
import { ITeamsMessage } from "../../../services/ITeamsMessage";

export interface IPnPGraphState {
    teams: ITeams[];
    channels: IChannels[];
    messages: ITeamsMessage[];
    selectedTeamsID: string;
    selectedChannelID: string;
    selectedChannel: string;
    hideDialog: boolean;
    message: string;
  }