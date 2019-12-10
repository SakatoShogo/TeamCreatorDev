import * as React from 'react';
import { ITeamCreatorProps } from './ITeamCreatorProps';
/**
 * 0.0.1
 * State of the component
 */
export declare enum CreationState {
    /**
     * Initial state - user input
     */
    notStarted = 0,
    /**
     * creating all selected elements (group, team, channel, tab)
     */
    creating = 1,
    /**
     * everything has been created
     */
    created = 2,
    /**
     * error during creation
     */
    error = 4
}
/**
 * 0.0.2
 * App definition returned from App Catalog
 */
export interface ITeamsApp {
    id: string;
    externalId?: string;
    displayName: string;
    version: string;
    distributionMethod: string;
}
/**
 * 0.0.3
 * TODO：どれが使われてて、どれが使われていないのか整理したい
 * State
 */
export interface ITeamCreatorState {
    /**
     * チーム名
     */
    teamName?: string;
    teamNickName?: string;
    /**
     * チーム説明
     */
    teamDescription?: string;
    /**
     * 所有者
     */
    owners?: string[];
    /**
     * メンバー
     */
    members?: string[];
    /**
     * Flag if channel should be created
     */
    createChannel?: boolean;
    /**
     * チャネル名
     */
    channelName?: string;
    /**
     * チャネル説明
     */
    channelDescription?: string;
    /**
     * flag if we need to add a tab
     */
    addTab?: boolean;
    /**
     * タブ名
     */
    tabName?: string;
    /**
     * teams apps from app catalog
     */
    apps?: ITeamsApp[];
    /**
     * current state of the component
     */
    creationState?: CreationState;
    /**
     * creation spinner text
     */
    spinnerText?: string;
    /**
     * id of the selected app to be added as tab
     */
    selectedAppId?: string;
    generalSelectedAppId?: string;
    addTabToGeneral?: boolean;
    generalTabName?: string;
    template?: string;
}
export default class TeamCreator extends React.Component<ITeamCreatorProps, ITeamCreatorState, {}> {
    constructor(props: ITeamCreatorProps);
    render(): React.ReactElement<ITeamCreatorProps>;
    private _onTeamNameChange;
    private _onTeamNickNameChange;
    private _onTeamDescriptionChange;
    private _onTemplateChange;
    private _onMembersSelected;
    private _onOwnersSelected;
    private _onCreateClick;
    private _onClearClick;
    private alertClicked;
    private _clearState;
    private _getAvailableApps;
    /**
     * Main flow
     */
    private _processCreationRequest;
    /**
     * 1.1 Creates O365 group
     * @param graphClient graph client
     */
    private _createGroup;
    /**
     * 1.2 Creates team. as mentioned in the documentation - we need to make multiple attempts if team creation request errored
     * @param groupId group id
     * @param graphClient graph client
     */
    private _createTeamWithAttempts;
    /**
     * Waits 10 seconds and tries to create a team
     * @param groupId group id
     * @param graphClient graph client
     */
    private _createTeam;
    /**
     * 1.3 Creates channel in the team
     * @param teamId team id
     * @param graphClient graph client
     */
    private _createGroupChannel;
    /**
     * 1.4 Installs the app to the team
     * @param teamId team Id
     * @param graphClient graph client
     */
    private _installApp;
    private _addGroupTab;
    private _onError;
    /**
     * 1.5 Adds tab to the specified channel of the team
     * @param teamId team id
     * @param channelId channel id
     * @param graphClient graph client
     */
    private _addGroupTab2;
    /**
     * Generates mail nick name by display name of the group
     * @param displayName group display name
     */
    private _generateMailNickname;
}
//# sourceMappingURL=TeamCreator.d.ts.map