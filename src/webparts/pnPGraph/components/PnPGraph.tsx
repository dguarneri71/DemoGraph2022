import * as React from 'react';
import styles from './PnPGraph.module.scss';
import { IPnPGraphProps } from './IPnPGraphProps';
import { escape } from '@microsoft/sp-lodash-subset';
import SPDataService from '../../../services/SPDataService';
import { IPnPGraphState } from './IPnPGraphState';
import { ActivityItem, Icon, Link, mergeStyleSets, Stack, IStackStyles, IStackTokens, IStackItemStyles, PrimaryButton, Panel, PanelType, TextField, Dialog, DialogType, getId } from 'office-ui-fabric-react';
import { stringIsNullOrEmpty } from "@pnp/core";

const classNames = mergeStyleSets({
  exampleRoot: {
    marginTop: '20px',
  },
  nameText: {
    fontWeight: 'bold',
  },
  popupMsg: {
    marginBottom: "10px"
  }
});

// Styles definition
const stackStyles: IStackStyles = {
  root: {
    background: 'white',
  },
};
const stackItemStyles: IStackItemStyles = {
  root: {
    alignItems: 'center',
    background: 'white',
    color: '#000',
    /*display: 'flex',*/
    /*height: 50,*/
    justifyContent: 'center',
  },
};

// Tokens definition
const stackTokens: IStackTokens = {
  childrenGap: 3
};

export default class PnPGraph extends React.Component<IPnPGraphProps, IPnPGraphState> {
  private graphService: SPDataService;
  private _labelId: string = getId('dialogLabel');
  private _subTextId: string = getId('subTextLabel');

  constructor(props: IPnPGraphProps, { }) {
    super(props);

    this.graphService = new SPDataService(this.props.context);

    this.state = {
      teams: [],
      channels: [],
      messages: [],
      selectedTeamsID: null,
      selectedChannelID: null,
      selectedChannel: null,
      hideDialog: true,
      message: null
    };
  }

  /**
   * Carico le informazioni
   */
  public componentWillMount(): void {
    this.graphService.getMyTeams().then(results => {
      console.log("my TEAMS: ", results);

      this.setState({
        teams: results
      });
    });
  }

  public render(): React.ReactElement<IPnPGraphProps> {
    const { hideDialog } = this.state;
    let teamItems = [];
    let channelItems = [];
    let msgItems = [];

    this.state.teams.forEach((team, index, teams) => {
      teamItems.push({
        key: index,
        activityDescription: [
          <Link
            key={1}
            className={classNames.nameText}
            onClick={() => {
              this.GetChannels(team["id"]);

              this.setState({
                selectedTeamsID: team["id"]
              });
            }}
          >
            {team["displayName"]}
          </Link>,
        ],
        activityIcon: <Icon iconName={'TeamsLogo'} />
      });
    });

    this.state.channels.forEach((channel, index, channels) => {
      channelItems.push({
        key: index,
        activityDescription: [
          <Link
            key={1}
            className={classNames.nameText}
            onClick={() => {
              this.GetMessage(this.state.selectedTeamsID, channel["id"]);

              this.setState({
                selectedChannelID: channel["id"],
                selectedChannel: channel["displayName"]
              });
            }}
          >
            {channel["displayName"]}
          </Link>,
        ],
        activityIcon: <Icon iconName={'CannedChat'} />
      });
    });

    this.state.messages.forEach((msg, index, messages) => {
      if (msg.messageType == "message" && msg.deletedDateTime == null) {
        msgItems.push({
          key: index,
          activityDescription: [
            <div>From: {msg.user}</div>,
            <div>Message: {msg.body.content}</div>
          ],
          activityIcon: <Icon iconName={'OfficeChat'} />
        });
      }
    });

    let disabled = stringIsNullOrEmpty(this.state.selectedTeamsID) || stringIsNullOrEmpty(this.state.selectedChannelID);

    return (
      <div className={styles.pnPGraph}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <Stack horizontal tokens={stackTokens}>
                <PrimaryButton text="Send message" onClick={this._showDialog.bind(this)} allowDisabledFocus disabled={disabled} />
              </Stack>
              <Stack horizontal styles={stackStyles} tokens={stackTokens} verticalAlign="start">
                <Stack.Item styles={stackItemStyles}>
                  <div><h3>Teams</h3></div>
                  <div>
                    {teamItems.map((item: { key: string | number }) => (
                      <ActivityItem {...item} key={item.key} className={classNames.exampleRoot} />
                    ))}
                  </div>
                </Stack.Item>
                <Stack.Item styles={stackItemStyles}>
                  <div><h3>Channels</h3></div>
                  <div>
                    {channelItems.map((item: { key: string | number }) => (
                      <ActivityItem {...item} key={item.key} className={classNames.exampleRoot} />
                    ))}
                  </div>
                </Stack.Item>
                <Stack.Item styles={stackItemStyles}>
                  <div><h3>Messages</h3></div>
                  <div>
                    {msgItems.map((item: { key: string | number }) => (
                      <ActivityItem {...item} key={item.key} className={classNames.exampleRoot} />
                    ))}
                  </div>
                </Stack.Item>
              </Stack>
            </div>
          </div>
        </div>
        <div>
          <Dialog
            hidden={hideDialog}
            onDismiss={this._closeDialog}
            dialogContentProps={{
              type: DialogType.normal,
              title: "Send a message",
              subText: null
            }}
            modalProps={{
              titleAriaId: this._labelId,
              subtitleAriaId: this._subTextId,
              isBlocking: false,
              styles: { main: { maxWidth: 450 } },
            }}
          >
            <div className={classNames.popupMsg}>
              <TextField label="Message" multiline rows={3} onChange={this._onChange} />
            </div>
            <div>
              <PrimaryButton text="Send" onClick={this._send.bind(this)} allowDisabledFocus />
            </div>
          </Dialog>
        </div>
      </div>
    );
  }

  private GetChannels(teamsID: string): void {
    this.graphService.getChannels(teamsID).then(results => {
      console.log("my CHANNELS: ", results);

      this.setState({
        channels: results
      });
    });
  }

  private GetMessage(teamsID: string, chennalsID: string): void {
    this.graphService.getMessages(teamsID, chennalsID).then(results => {
      console.log("my MESSAGES: ", results);

      this.setState({
        messages: results
      });
    });
  }

  private SendMessage(teamsID: string, chennalsID: string, msg: string): void {
    this.graphService.sendMessage(teamsID, chennalsID, msg).then(result=>{
      console.log("SendMessage() - result: ", result);
      this.GetMessage(teamsID, chennalsID);
    });

    this.setState({
      hideDialog: true
    });
  }

  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

  private _onChange = (ev: any, newText: string): void => {
    this.setState({ message: newText });
  }

  private _send = (): void => {
    const { selectedChannelID, selectedTeamsID, message } = this.state;
    this.SendMessage(selectedTeamsID, selectedChannelID, message);
  }
}