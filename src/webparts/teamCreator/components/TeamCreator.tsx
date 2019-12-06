import * as React from 'react';
import styles from './TeamCreator.module.scss';
import { ITeamCreatorProps } from './ITeamCreatorProps';

import { Icon, IIconStyles, Image, Stack, IStackTokens, Text, ITextStyles } from 'office-ui-fabric-react';
import { TextField, ITextField } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { PrimaryButton, DefaultButton, ActionButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as strings from 'TeamCreatorWebPartStrings';
import { MSGraphClient } from '@microsoft/sp-http';
import { autobind } from '@uifabric/utilities';
import { string } from 'prop-types';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import { IStyleSet } from 'office-ui-fabric-react/lib/Styling';
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
// TODO: move this file to Accordion
import { Accordion } from '@uifabric/experiments/lib/Accordion';

import { Card, ICardTokens, ICardSectionStyles, ICardSectionTokens } from '@uifabric/react-cards';
import { FontWeights } from '@uifabric/styling';

const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
  root: { marginTop: 10 }
};
/**
 * State of the component
 */
export enum CreationState {
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
 * State
 */
export interface ITeamCreatorState {
  /**
   * Selected team name. Also used as group name
   */
  teamName?: string;
  teamNickName?: string;
  /**
   * team description
   */
  teamDescription?: string;
  /**
   * Group owners
   */
  owners?: string[];
  /**
   * group members
   */
  members?: string[];
  /**
   * Flag if channel should be created
   */
  createChannel?: boolean;
  /**
   * channel name
   */
  channelName?: string;
  /**
   * channel description
   */
  channelDescription?: string;
  /**
   * flag if we need to add a tab
   */
  addTab?: boolean;
  /**
   * tab name
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
  template?: string ;
}

export default class TeamCreator extends React.Component<ITeamCreatorProps, ITeamCreatorState,{}> {

  constructor(props: ITeamCreatorProps) {
    super(props);

    this.state = {
      creationState: CreationState.notStarted
    };

    this._onClearClick = this._onClearClick.bind(this);
  }

  public render(): React.ReactElement<ITeamCreatorProps> {

    const {
      teamName,
      teamNickName,
      teamDescription,
      createChannel,
      channelName,
      channelDescription,
      addTab,
      tabName,
      apps,
      creationState,
      spinnerText,
      selectedAppId,
      generalSelectedAppId,
      addTabToGeneral,
      generalTabName,
      template
    } = this.state;

    const appsDropdownOptions: IDropdownOption[] = apps ? apps.map(app => { return { key: app.id, text: app.displayName }; }) : [];
    
    const alertClicked = (): void => {
      alert('Clicked');
    };

    const siteTextStyles: ITextStyles = {
      root: {
        color: '#025F52',
        fontWeight: FontWeights.semibold
      }
    };
    const descriptionTextStyles: ITextStyles = {
      root: {
        color: '#333333',
        fontWeight: FontWeights.regular
      }
    };
    const helpfulTextStyles: ITextStyles = {
      root: {
        color: '#333333',
        fontWeight: FontWeights.regular
      }
    };
    const iconStyles: IIconStyles = {
      root: {
        color: '#0078D4',
        fontSize: 16,
        fontWeight: FontWeights.regular
      }
    };
    const footerCardSectionStyles: ICardSectionStyles = {
      root: {
        borderLeft: '1px solid #F3F2F1'
      }
    };

    const sectionStackTokens: IStackTokens = { childrenGap: 20};
    const cardTokens: ICardTokens = { childrenMargin: 12,minWidth:'700px'};
    const footerCardSectionTokens: ICardSectionTokens = { padding: '0px 0px 0px 12px'};

    return (
      <div className={styles.teamCreator}>
       
        <div className={styles.container}>
          {{
            0: 
            <div>
                <Pivot>
                  <PivotItem
                    headerText="トップ"
                    headerButtonProps={{
                      'data-order': 1,
                      'data-title': 'Application'
                    }}
                    itemIcon="DietPlanNotebook"
                  >

                    
                    {/** タイトル */}
                    <h2 className={styles.sectionTitle}>{strings.Welcome}</h2>
                    <div className={styles.teamSection}>
                      {/** チーム名 */}
                      <TextField required={true} label={strings.TeamNameLabel} value={teamName} onChanged={this._onTeamNameChange.bind(this)}></TextField>

                      {/** チームアドレス */}
                      <TextField required={true} label={strings.TeamNickNameLabel} value={teamNickName} suffix="@rjtk1114.onmicrosoft.com" onChanged={this._onTeamNickNameChange.bind(this)}></TextField>

                      {/** チーム説明 */}
                      <TextField label={strings.TeamDescriptionLabel} value={teamDescription} onChanged={this._onTeamDescriptionChange.bind(this)}></TextField>

                      {/** 所有者 */}
                      <PeoplePicker
                        context={this.props.context}
                        titleText={strings.Owners}
                        personSelectionLimit={3}
                        showHiddenInUI={false}
                        selectedItems={this._onOwnersSelected.bind(this)}
                        isRequired={true} />

                      {/** メンバー */}
                      <PeoplePicker
                        context={this.props.context}
                        titleText={strings.Members}
                        personSelectionLimit={3}
                        showHiddenInUI={false}
                        selectedItems={this._onMembersSelected.bind(this)} />
                    </div>
                    <div>
                      <ChoiceGroup
                        onChange={this._onTemplateChange.bind(this)}
                        label="テンプレート"
                        defaultSelectedKey='group'
                        
                        options={[
                          {
                            key: 'group',
                            imageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/group.png',
                            selectedImageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/group.png',
                            text: '部門',
                            style:{
                              marginRight: '100px'
                            }
                          },
                          {
                            key: 'project',
                            imageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/project.png',
                            selectedImageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/project.png',
                            text: 'プロジェクト'
                          },
                          {
                            key: 'child',
                            imageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/child.png',
                            selectedImageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/child.png',
                            text: '新入社員'
                          },
                          {
                            key: 'custom',
                            imageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/edit.png',
                            selectedImageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/edit.png',
                            text: '研修'
                          }
                        ]}
                      /></div>


                    {template == 'custom' &&
                      <div>

                        <div className={styles.channelSection}>
                          <TextField required={createChannel} label={strings.ChannelName} value={channelName} onChanged={this._onChannelNameChange.bind(this)}></TextField>
                          <TextField label={strings.ChannelDescription} value={channelDescription} onChanged={this._onChannelDescriptionChange.bind(this)}></TextField>
                        </div>
                        <Checkbox label={strings.AddTab} checked={addTab} onChange={this._onAddTabChange.bind(this)} />
                        {addTab && <div>
                          <TextField required={addTab} label={strings.TabName} value={tabName} onChanged={this._onTabNameChange.bind(this)}></TextField>
                          <Dropdown
                            required={addTab}
                            label={strings.App}
                            disabled={!this.state.apps}
                            options={appsDropdownOptions}
                            selectedKey={selectedAppId}
                            onChanged={this._onAppSelected.bind(this)}></Dropdown>

                        </div>}
                      </div>
                    }

                    <div className={styles.buttons}>
                      {/** 申請ボタン */}
                      <PrimaryButton text={strings.Create} className={styles.button} onClick={this._onCreateClick.bind(this)} />

                      {/** クリアボタン */}
                      <DefaultButton text={strings.Clear} className={styles.button} onClick={this._onClearClick} />
                    </div>
                  </PivotItem>

                  <PivotItem headerText="設定" itemIcon='Settings'>
                   
                    <h2 className={styles.sectionTitle}>テンプレート一覧</h2>
                    <div>
                      <Stack tokens={sectionStackTokens}>
                        <Card horizontal onClick={alertClicked} tokens={cardTokens}>
                          <Card.Item fill>
                            <Image src="https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/group_tem.png" alt="Placeholder image." />
                          </Card.Item>
                          <Card.Section>
                            <Text variant="small" styles={siteTextStyles}>部門</Text>
                            <Text styles={descriptionTextStyles}>部門向けのチームテンプレートになります</Text>
                          </Card.Section>
                        </Card>

                        <Card horizontal onClick={alertClicked} tokens={cardTokens}>
                          <Card.Item fill>
                            <Image src="https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/project_tem.png" alt="Placeholder image." />
                          </Card.Item>
                          <Card.Section>
                            <Text variant="small" styles={siteTextStyles}>プロジェクト</Text>
                            <Text styles={descriptionTextStyles}>プロジェクト向けのチームテンプレートになります</Text>
                          </Card.Section>
                        </Card>

                        <Card horizontal onClick={alertClicked} tokens={cardTokens}>
                          <Card.Item fill>
                            <Image src="https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/child_tem.png" alt="Placeholder image." />
                          </Card.Item>
                          <Card.Section>
                            <Text variant="small" styles={siteTextStyles}>新入社員</Text>
                            <Text styles={descriptionTextStyles}>新入社員向けのチームテンプレートになります</Text>
                          </Card.Section>
                        </Card>
                      </Stack>
                    </div>

                  </PivotItem>
                </Pivot>

            </div>,
            1: <div>
              <Spinner label={spinnerText} />
            </div>,
            2: <div>
              <h2 className={styles.success} >{strings.Success}</h2>
              <PrimaryButton className={styles.goTeams} iconProps={{ iconName: 'TeamsLogo' }} href='https://aka.ms/mstfw' target='_blank'>{strings.OpenTeams}</PrimaryButton>
              <DefaultButton onClick={this._onClearClick}>{strings.StartOver}</DefaultButton>
            </div>,
            4: <div>
              <div className={styles.error}>{strings.Error}</div>
              <DefaultButton onClick={this._onClearClick}>{strings.StartOver}</DefaultButton>
            </div>
          }[creationState]}
        </div>
      </div>
    );
  }
  //  チーム名
  private _onTeamNameChange(value: string) {
    this.setState({
      teamName: value
    });
  }
  private _onTeamNickNameChange(value: string) {
    this.setState({
      teamNickName: value
    });
  }
  //  チーム説明
  private _onTeamDescriptionChange(value: string) {
    this.setState({
      teamDescription: value
    });
  }
  
  //  チャネル名
  private _onChannelNameChange(value: string) {
    this.setState({
      channelName: value
    });
  }

  //  チャネル説明
  private _onGeneralTabNameChange(value: string) {
    this.setState({
      generalTabName: value
    });
  }

  //  チャネル説明
  private _onChannelDescriptionChange(value: string) {
    this.setState({
      channelDescription: value
    });
  }

  //  タブ名
  private _onTabNameChange(value: string) {
    this.setState({
      tabName: value
    });
  }

  //  アプリの選択
  private _onAppSelected(item: IDropdownOption) {
    this.setState({
      selectedAppId: item.key as string
    });
  }

  //  アプリの選択
  private _onGeneralAppSelected(item: IDropdownOption) {
    this.setState({
      generalSelectedAppId: item.key as string
    });
  }
  
  //  テンプレート
  private _onTemplateChange(e: React.FormEvent<HTMLElement | HTMLInputElement>, option: IChoiceGroupOption) {
    const optionKey = option.key;

    this.setState({
      template: optionKey
    });
  }

  //  チャネルの追加
  private _onCreateChannelChange(e: React.FormEvent<HTMLElement | HTMLInputElement>, checked: boolean) {
    this.setState({
      createChannel: checked
    });
  }

  //  タブの追加
  private _onAddTabChange(e: React.FormEvent<HTMLElement | HTMLInputElement>, checked: boolean) {
    this.setState({
      addTab: checked
    });

    this._getAvailableApps();
  }

  //  タブの追加
  private _onAddTabToGeneralChange(e: React.FormEvent<HTMLElement | HTMLInputElement>, checked: boolean) {
    this.setState({
      addTabToGeneral: checked
    });

    this._getAvailableApps();
  }

  //　メンバー
  private _onMembersSelected(members: IPeoplePickerUserItem[]) {
    console.log(members);
    
    this.setState({
      members: members.map(m => m.id)
    });
  }

  //  所有者
  private _onOwnersSelected(owners: IPeoplePickerUserItem[]) {
    this.setState({
      owners: owners.map(o => o.id)
    });
  }

  //  申請
  private async _onCreateClick() {
    this._processCreationRequest();
  }

  //  キャンセル
  private _onClearClick() {
    this._clearState();
  }

  private _clearState() {
    this.setState({
      teamName: '',
      teamDescription: '',
      members: [],
      owners: [],
      createChannel: false,
      channelName: '',
      channelDescription: '',
      addTab: false,
      tabName: '',
      selectedAppId: '',
      generalSelectedAppId:'',
      creationState: CreationState.notStarted,
      spinnerText: ''
    });
  }

  //  利用可能なアプリ取得
  private async _getAvailableApps(): Promise<void> {
    if (this.state.apps) {
      return;
    }

    const context = this.props.context;
    const graphClient = await context.msGraphClientFactory.getClient();

    const appsResponse = await graphClient.api('appCatalogs/teamsApps').version('v1.0').get();
    const apps = appsResponse.value as ITeamsApp[];
    apps.sort((a, b) => {
      if (a.displayName < b.displayName) {
        return -1;
      }
      else if (a.displayName > b.displayName) {
        return 1;
      }
      return 0;
    });

    this.setState({
      apps: apps
    });
  }

  /**
   * Main flow
   */
  private async _processCreationRequest(): Promise<void> {
    const context = this.props.context;
    // Graph Client の初期化
    const graphClient = await context.msGraphClientFactory.getClient();

    this.setState({
      creationState: CreationState.creating,
      spinnerText: strings.CreatingGroup
    });

    //
    // Office365 グループ作成
    //
    const groupId = await this._createGroup(graphClient);
    
    if (!groupId) {
      this._onError();
      return;
    }
    console.log(groupId);

    this.setState({
      spinnerText: strings.CreatingTeam
    });

    //
    // チーム作成
    //
    const teamId = await this._createTeamWithAttempts(groupId, graphClient);
    if (!teamId) {
      this._onError();
      return;
    }

    if (this.state.template == 'group') {
      
      this._getAvailableApps();

      const channelAId = await this._createGroupChannel(teamId, '01.Aグループ', 'Aグループ用です。', graphClient);
      if (!channelAId) {
        this._onError();
        return;
      }

      const channelBId = await this._createGroupChannel(teamId, '02.Bグループ', 'Bグループ用です。', graphClient);
      if (!channelBId) {
        this._onError();
        return;
      }

      const channelCId = await this._createGroupChannel(teamId, '03.Cグループ', 'Cグループ用です。', graphClient);
      if (!channelCId) {
        this._onError();
        return;
      }

      const channelDId = await this._createGroupChannel(teamId, '04.Dグループ', 'Dグループ用です。', graphClient);
      if (!channelDId) {
        this._onError();
        return;
      }

      const channelEId = await this._createGroupChannel(teamId, '05.Eグループ', 'Eグループ用です。', graphClient);
      if (!channelEId) {
        this._onError();
        return;
      }

      // アプリインストール
      const SPAPP = this.state.apps.filter(element => element.displayName == 'SharePoint');
      const OneNote = this.state.apps.filter(element => element.displayName == 'OneNote');

      const isSPInstalled = await this._installApp(teamId, SPAPP[0].id, graphClient);
      if (!isSPInstalled) {
        this._onError();
        return;
      }

      const isOneNoteInstalled = await this._installApp(teamId, OneNote[0].id, graphClient);
      if (!isOneNoteInstalled) {
        this._onError();
        return;
      }

      //
      // タブ追加
      //
      this._addGroupTab(teamId, channelAId, 'SharePoint',SPAPP[0].id, graphClient);
      this._addGroupTab(teamId, channelAId, 'OneNote',OneNote[0].id, graphClient);

      this._addGroupTab(teamId, channelBId, 'SharePoint', SPAPP[0].id, graphClient);
      this._addGroupTab(teamId, channelBId, 'OneNote', OneNote[0].id, graphClient);

      this._addGroupTab(teamId, channelCId, 'SharePoint', SPAPP[0].id, graphClient);
      this._addGroupTab(teamId, channelCId, 'OneNote', OneNote[0].id, graphClient);
      
      this._addGroupTab(teamId, channelDId, 'SharePoint', SPAPP[0].id, graphClient);
      this._addGroupTab(teamId, channelDId, 'OneNote', OneNote[0].id, graphClient);

      this._addGroupTab(teamId, channelEId, 'SharePoint', SPAPP[0].id, graphClient);
      this._addGroupTab(teamId, channelEId, 'OneNote', OneNote[0].id, graphClient);

      this.setState({
        spinnerText: strings.CreatingTab
      });

    } else if (this.state.template == 'project') {

    } else {
      //
      //  一般チャネルを取得
      //
      //const generalChannelId = this._getChannel(teamId,graphClient);
      //if (!generalChannelId){
      //  this._onError();
      //  return;
      //}

      //
      //  一般チャネルにタブ追加
      //
      //const isInstalledToGeneral = await this._installApp(teamId, this.state.generalSelectedAppId,graphClient);
      //if (!isInstalledToGeneral) {
      //  this._onError();
      //  return;

      //}
      //const isTabCreatedToGeneral = await this._addTab(teamId, await generalChannelId, this.state.generalSelectedAppId, graphClient);
      //if (!isTabCreatedToGeneral) {
      //  this._onError();
      //  return;
      //}

      //if (!this.state.createChannel) {
      this.setState({
        creationState: CreationState.created
      });
      //return;
      //}

      this.setState({
        spinnerText: strings.CreatingChannel
      });

      //
      // チャネル作成
      //
      const channelId = await this._createChannel(teamId, graphClient);
      if (!channelId) {
        this._onError();
        return;
      }
      // 
      //  タブ追加にチェックが入っていれば次に進む
      //
      if (!this.state.addTab) {
        this.setState({
          creationState: CreationState.created
        });
        return;
      }

      this.setState({
        spinnerText: strings.InstallingApp
      });

      //
      // アプリインストール
      //
      const isInstalled = await this._installApp(teamId, this.state.selectedAppId, graphClient);
      if (!isInstalled) {
        this._onError();
        return;
      }

      this.setState({
        spinnerText: strings.CreatingTab
      });

      //
      // タブ追加
      //
      const isTabCreated = await this._addTab(teamId, channelId, this.state.selectedAppId, graphClient);
      if (!isTabCreated) {
        this._onError();
      }
      else {
        this.setState({
          creationState: CreationState.created
        });
      }
    }
  }

  private async _addGroupTab(teamId: string, channelId: string,tabName:string ,appId: string, graphClient: MSGraphClient):Promise<void>{

    //
    // タブ追加
    //
    const isTabCreated = await this._addGroupTab2(teamId, channelId, tabName,appId, graphClient);
    if (!isTabCreated) {
      this._onError();
    }
    else {
      this.setState({
        creationState: CreationState.created
      });
    }
  }

  private _onError(message?: string): void {
    this.setState({
      creationState: CreationState.error
    });
  }

  /**
   * Installs the app to the team
   * @param teamId team Id
   * @param graphClient graph client
   */
  private async _installApp(teamId: string, selectedAppId:string , graphClient: MSGraphClient): Promise<boolean> {
    try {
      await graphClient.api(`teams/${teamId}/installedApps`).version('v1.0').post({
        'teamsApp@odata.bind': `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${selectedAppId}`
      });
    }
    catch (error) {
      console.error(error);
      return false;
    }

    return true;
  }

  /**
   * Adds tab to the specified channel of the team
   * @param teamId team id
   * @param channelId channel id
   * @param graphClient graph client
   */
  private async _addTab(teamId: string, channelId: string, selectedAppId:string,graphClient: MSGraphClient): Promise<boolean> {
    try {
      await graphClient.api(`teams/${teamId}/channels/${channelId}/tabs`).version('v1.0').post({
        displayName: this.state.tabName,
        'teamsApp@odata.bind': `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${selectedAppId}`
      });
    }
    catch (error) {
      console.error(error);
      return false;
    }

    return true;
  }

  /**
   * Adds tab to the specified channel of the team
   * @param teamId team id
   * @param channelId channel id
   * @param graphClient graph client
   */
  private async _addGroupTab2(teamId: string, channelId: string, tabName:string,selectedAppId: string, graphClient: MSGraphClient): Promise<boolean> {
    try {
      await graphClient.api(`teams/${teamId}/channels/${channelId}/tabs`).version('v1.0').post({
        displayName: tabName,
        'teamsApp@odata.bind': `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${selectedAppId}`
      });
    }
    catch (error) {
      console.error(error);
      return false;
    }

    return true;
  }

  /**
   * Creates channel in the team
   * @param teamId team id
   * @param graphClient graph client
   */
  private async _createChannel(teamId: string, graphClient: MSGraphClient): Promise<string> {
    const {
      channelName,
      channelDescription
    } = this.state;

    try {
      const response = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').post({
        displayName: channelName,
        description: channelDescription
      });

      return response.id;
    }
    catch (error) {
      console.error(error);
      return '';
    }
  }

  /**
   * Creates channel in the team
   * @param teamId team id
   * @param graphClient graph client
   */
  private async _createGroupChannel(teamId: string, channelName: string, channelDescription:string,graphClient: MSGraphClient): Promise<string> {
    try {
      const response = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').post({
        displayName: channelName,
        description: channelDescription
      });

      return response.id;
    }
    catch (error) {
      console.error(error);
      return '';
    }
  }

  /**
 * Creates channel in the team
 * @param teamId team id
 * @param graphClient graph client
 */
  private async _getChannel(teamId: string, graphClient: MSGraphClient): Promise<string> {

    try {
      const response = await graphClient.api(`teams/${teamId}/channels`).version('v1.0').get();
      console.log(response);
      return response.id;
    }
    catch (error) {
      console.error(error);
      return '';
    }
  }

  /**
   * Creates O365 group
   * @param graphClient graph client
   */
  private async _createGroup(graphClient: MSGraphClient): Promise<string> {
    const displayName = this.state.teamName;
    const mailNickname = this._generateMailNickname(this.state.teamNickName);

    let {
      owners,
      members
    } = this.state;

    const groupRequest = {
      displayName: displayName,
      description: this.state.teamDescription,
      groupTypes: [
        'Unified'
      ],
      mailEnabled: true,
      mailNickname: mailNickname,
      securityEnabled: false
    };
    if (owners && owners.length) {
      groupRequest['owners@data.bind'] = owners.map(owner => {
        console.log(owner);
        return `https://graph.microsoft.com/v1.0/users/${owner}`;
      });
    }
    if (members && members.length) {
      groupRequest['members@data.bind'] = members.map(member => {
        console.log(member);
        return `https://graph.microsoft.com/v1.0/users/${member}`;
      });
    }

    console.log(groupRequest);

    try {
      const response = await graphClient.api('groups').version('v1.0').post(groupRequest);
      console.log(response);
      return response.id;
    }
    catch (error) {
      console.error(error);
      return '';
    }
  }

  /**
   * Creates team. as mentioned in the documentation - we need to make multiple attempts if team creation request errored
   * @param groupId group id
   * @param graphClient graph client
   */
  private async _createTeamWithAttempts(groupId: string, graphClient: MSGraphClient): Promise<string> {

    let attemptsCount = 0;
    let teamId: string = '';

    //
    // From the documentation: If the group was created less than 15 minutes ago, it's possible for the Create team call to fail with a 404 error code due to replication delays. 
    // The recommended pattern is to retry the Create team call three times, with a 10 second delay between calls.
    //
    do {
      teamId = await this._createTeam(groupId, graphClient);
      if (teamId) {
        attemptsCount = 3;
      }
      else {
        attemptsCount++;
      }
    } while (attemptsCount < 3);

    return teamId;
  }

  /**
   * Waits 10 seconds and tries to create a team
   * @param groupId group id
   * @param graphClient graph client
   */
  private async _createTeam(groupId: string, graphClient: MSGraphClient): Promise<string> {
    return new Promise<string>(resolve => {
      setTimeout(() => {
        graphClient.api(`groups/${groupId}/team`).version('v1.0').put({
          memberSettings: {
            allowCreateUpdateChannels: true
          },
          messagingSettings: {
            allowUserEditMessages: true,
            allowUserDeleteMessages: true
          },
          funSettings: {
            allowGiphy: true,
            giphyContentRating: "strict"
          }
        }).then(response => {
          resolve(response.id);
        }, () => {
          resolve('');
        });
      }, 10000);
    });
  }

  /**
   * Generates mail nick name by display name of the group
   * @param displayName group display name
   */
  private _generateMailNickname(displayName: string): string {
    return displayName.toLowerCase().replace(/\s/gmi, '-');
  }
}
