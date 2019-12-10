var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
// SCSS
import styles from './TeamCreator.module.scss';
// String
import * as strings from 'TeamCreatorWebPartStrings';
// Office UI Fabric
import { Image, Stack, Text } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { ChoiceGroup } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import { Card } from '@uifabric/react-cards';
import { FontWeights } from '@uifabric/styling';
// PnP PeoplePicker
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
/**
 * 0.0.1
 * State of the component
 */
export var CreationState;
(function (CreationState) {
    /**
     * Initial state - user input
     */
    CreationState[CreationState["notStarted"] = 0] = "notStarted";
    /**
     * creating all selected elements (group, team, channel, tab)
     */
    CreationState[CreationState["creating"] = 1] = "creating";
    /**
     * everything has been created
     */
    CreationState[CreationState["created"] = 2] = "created";
    /**
     * error during creation
     */
    CreationState[CreationState["error"] = 4] = "error";
})(CreationState || (CreationState = {}));
var TeamCreator = /** @class */ (function (_super) {
    __extends(TeamCreator, _super);
    function TeamCreator(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            creationState: CreationState.notStarted
        };
        _this._onClearClick = _this._onClearClick.bind(_this);
        return _this;
    }
    TeamCreator.prototype.render = function () {
        // State
        var _a = this.state, teamName = _a.teamName, teamNickName = _a.teamNickName, teamDescription = _a.teamDescription, apps = _a.apps, creationState = _a.creationState, spinnerText = _a.spinnerText;
        // アプリ一覧
        var appsDropdownOptions = apps ? apps.map(function (app) { return { key: app.id, text: app.displayName }; }) : [];
        // Styles of Card 
        // テンプレートタイトル
        var siteTextStyles = {
            root: {
                color: '#025F52',
                fontWeight: FontWeights.semibold
            }
        };
        // テンプレート説明
        var descriptionTextStyles = {
            root: {
                color: '#333333',
                fontWeight: FontWeights.regular
            }
        };
        var sectionStackTokens = {
            childrenGap: 20
        };
        var cardTokens = {
            childrenMargin: 12,
            minWidth: '700px'
        };
        // レンダリング
        return (React.createElement("div", { className: styles.teamCreator },
            React.createElement("div", { className: styles.container }, {
                0: React.createElement("div", null,
                    React.createElement(Pivot, null,
                        React.createElement(PivotItem, { headerText: "\u30C8\u30C3\u30D7", headerButtonProps: {
                                'data-order': 1,
                                'data-title': 'Application'
                            }, itemIcon: "DietPlanNotebook" },
                            React.createElement("h2", { className: styles.sectionTitle }, strings.Welcome),
                            React.createElement("div", { className: styles.teamSection },
                                React.createElement(TextField, { required: true, label: strings.TeamNameLabel, value: teamName, onChanged: this._onTeamNameChange.bind(this) }),
                                React.createElement(TextField, { required: true, label: strings.TeamNickNameLabel, value: teamNickName, suffix: "@xxx.onmicrosoft.com", onChanged: this._onTeamNickNameChange.bind(this) }),
                                React.createElement(TextField, { label: strings.TeamDescriptionLabel, value: teamDescription, onChanged: this._onTeamDescriptionChange.bind(this) }),
                                React.createElement(PeoplePicker, { context: this.props.context, titleText: strings.Owners, personSelectionLimit: 3, showHiddenInUI: false, selectedItems: this._onOwnersSelected.bind(this), isRequired: true }),
                                React.createElement(PeoplePicker, { context: this.props.context, titleText: strings.Members, personSelectionLimit: 3, showHiddenInUI: false, selectedItems: this._onMembersSelected.bind(this) }),
                                React.createElement(ChoiceGroup, { onChange: this._onTemplateChange.bind(this), label: "\u30C6\u30F3\u30D7\u30EC\u30FC\u30C8", defaultSelectedKey: 'Department', options: [
                                        {
                                            key: 'Department',
                                            imageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/group.png',
                                            selectedImageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/group.png',
                                            text: strings.Department
                                        },
                                        {
                                            key: 'Project',
                                            imageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/project.png',
                                            selectedImageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/project.png',
                                            text: strings.Project
                                        },
                                        {
                                            key: 'NewEmployee',
                                            imageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/child.png',
                                            selectedImageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/child.png',
                                            text: strings.NewEmployee
                                        },
                                        {
                                            key: 'Training',
                                            imageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/training.png',
                                            selectedImageSrc: 'https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/training.png',
                                            text: strings.Training
                                        }
                                    ] })),
                            React.createElement("div", { className: styles.buttons },
                                React.createElement(PrimaryButton, { text: strings.Create, className: styles.button, onClick: this._onCreateClick.bind(this) }),
                                React.createElement(DefaultButton, { text: strings.Clear, className: styles.button, onClick: this._onClearClick }))),
                        React.createElement(PivotItem, { headerText: "\u8A2D\u5B9A", itemIcon: 'Settings' },
                            React.createElement("h2", { className: styles.sectionTitle }, "\u30C6\u30F3\u30D7\u30EC\u30FC\u30C8\u4E00\u89A7"),
                            React.createElement("div", null,
                                React.createElement(Stack, { tokens: sectionStackTokens },
                                    React.createElement(Card, { horizontal: true, onClick: this.alertClicked, tokens: cardTokens },
                                        React.createElement(Card.Item, { fill: true },
                                            React.createElement(Image, { src: "https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/group_tem.png", alt: "Placeholder image." })),
                                        React.createElement(Card.Section, null,
                                            React.createElement(Text, { variant: "small", styles: siteTextStyles }, strings.Department),
                                            React.createElement(Text, { styles: descriptionTextStyles }, strings.DepartmentDesc))),
                                    React.createElement(Card, { horizontal: true, onClick: this.alertClicked, tokens: cardTokens },
                                        React.createElement(Card.Item, { fill: true },
                                            React.createElement(Image, { src: "https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/project_tem.png", alt: "Placeholder image." })),
                                        React.createElement(Card.Section, null,
                                            React.createElement(Text, { variant: "small", styles: siteTextStyles }, strings.Project),
                                            React.createElement(Text, { styles: descriptionTextStyles }, strings.ProjectDesc))),
                                    React.createElement(Card, { horizontal: true, onClick: this.alertClicked, tokens: cardTokens },
                                        React.createElement(Card.Item, { fill: true },
                                            React.createElement(Image, { src: "https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/child_tem.png", alt: "Placeholder image." })),
                                        React.createElement(Card.Section, null,
                                            React.createElement(Text, { variant: "small", styles: siteTextStyles }, strings.NewEmployee),
                                            React.createElement(Text, { styles: descriptionTextStyles }, strings.NewEmployeeDesc))),
                                    React.createElement(Card, { horizontal: true, onClick: this.alertClicked, tokens: cardTokens },
                                        React.createElement(Card.Item, { fill: true },
                                            React.createElement(Image, { src: "https://rjtk1114.sharepoint.com/sites/SAMPLE001/SiteAssets/training_tem.png", alt: "Placeholder image." })),
                                        React.createElement(Card.Section, null,
                                            React.createElement(Text, { variant: "small", styles: siteTextStyles }, strings.Training),
                                            React.createElement(Text, { styles: descriptionTextStyles }, strings.TrainingDesc)))))))),
                1: React.createElement("div", null,
                    React.createElement(Spinner, { label: spinnerText })),
                2: React.createElement("div", null,
                    React.createElement("h2", { className: styles.success }, strings.Success),
                    React.createElement(PrimaryButton, { className: styles.goTeams, iconProps: { iconName: 'TeamsLogo' }, href: 'https://aka.ms/mstfw', target: '_blank' }, strings.OpenTeams),
                    React.createElement(DefaultButton, { onClick: this._onClearClick }, strings.StartOver)),
                4: React.createElement("div", null,
                    React.createElement("div", { className: styles.error }, strings.Error),
                    React.createElement(DefaultButton, { onClick: this._onClearClick }, strings.StartOver))
            }[creationState])));
    };
    //  チーム名
    TeamCreator.prototype._onTeamNameChange = function (value) {
        this.setState({
            teamName: value
        });
    };
    //  チームアドレス
    TeamCreator.prototype._onTeamNickNameChange = function (value) {
        this.setState({
            teamNickName: value
        });
    };
    //  チーム説明
    TeamCreator.prototype._onTeamDescriptionChange = function (value) {
        this.setState({
            teamDescription: value
        });
    };
    //  テンプレート
    TeamCreator.prototype._onTemplateChange = function (e, option) {
        var optionKey = option.key;
        this.setState({
            template: optionKey
        });
    };
    //　メンバー
    TeamCreator.prototype._onMembersSelected = function (members) {
        console.log(members);
        this.setState({
            members: members.map(function (m) { return m.id; })
        });
    };
    //  所有者
    TeamCreator.prototype._onOwnersSelected = function (owners) {
        this.setState({
            owners: owners.map(function (o) { return o.id; })
        });
    };
    //  申請
    TeamCreator.prototype._onCreateClick = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                this._processCreationRequest();
                return [2 /*return*/];
            });
        });
    };
    //  キャンセル
    TeamCreator.prototype._onClearClick = function () {
        this._clearState();
    };
    //  テンプレート一覧クリック
    TeamCreator.prototype.alertClicked = function () {
    };
    TeamCreator.prototype._clearState = function () {
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
            generalSelectedAppId: '',
            creationState: CreationState.notStarted,
            spinnerText: ''
        });
    };
    //  利用可能なアプリ取得
    TeamCreator.prototype._getAvailableApps = function () {
        return __awaiter(this, void 0, void 0, function () {
            var context, graphClient, appsResponse, apps;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.state.apps) {
                            return [2 /*return*/];
                        }
                        context = this.props.context;
                        return [4 /*yield*/, context.msGraphClientFactory.getClient()];
                    case 1:
                        graphClient = _a.sent();
                        return [4 /*yield*/, graphClient.api('appCatalogs/teamsApps').version('v1.0').get()];
                    case 2:
                        appsResponse = _a.sent();
                        apps = appsResponse.value;
                        apps.sort(function (a, b) {
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
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Main flow
     */
    TeamCreator.prototype._processCreationRequest = function () {
        return __awaiter(this, void 0, void 0, function () {
            var context, graphClient, groupId, teamId, channelAId, channelBId, channelCId, channelDId, channelEId, SPAPP, OneNote, isSPInstalled, isOneNoteInstalled;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        context = this.props.context;
                        this.setState({
                            creationState: CreationState.creating,
                            spinnerText: strings.CreatingGroup
                        });
                        return [4 /*yield*/, context.msGraphClientFactory.getClient()];
                    case 1:
                        graphClient = _a.sent();
                        return [4 /*yield*/, this._createGroup(graphClient)];
                    case 2:
                        groupId = _a.sent();
                        if (!groupId) {
                            this._onError();
                            return [2 /*return*/];
                        }
                        this.setState({
                            spinnerText: strings.CreatingTeam
                        });
                        return [4 /*yield*/, this._createTeamWithAttempts(groupId, graphClient)];
                    case 3:
                        teamId = _a.sent();
                        if (!teamId) {
                            this._onError();
                            return [2 /*return*/];
                        }
                        if (!(this.state.template == 'group')) return [3 /*break*/, 11];
                        this._getAvailableApps();
                        return [4 /*yield*/, this._createGroupChannel(teamId, '01.Aグループ', 'Aグループ用です。', graphClient)];
                    case 4:
                        channelAId = _a.sent();
                        if (!channelAId) {
                            this._onError();
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this._createGroupChannel(teamId, '02.Bグループ', 'Bグループ用です。', graphClient)];
                    case 5:
                        channelBId = _a.sent();
                        if (!channelBId) {
                            this._onError();
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this._createGroupChannel(teamId, '03.Cグループ', 'Cグループ用です。', graphClient)];
                    case 6:
                        channelCId = _a.sent();
                        if (!channelCId) {
                            this._onError();
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this._createGroupChannel(teamId, '04.Dグループ', 'Dグループ用です。', graphClient)];
                    case 7:
                        channelDId = _a.sent();
                        if (!channelDId) {
                            this._onError();
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this._createGroupChannel(teamId, '05.Eグループ', 'Eグループ用です。', graphClient)];
                    case 8:
                        channelEId = _a.sent();
                        if (!channelEId) {
                            this._onError();
                            return [2 /*return*/];
                        }
                        SPAPP = this.state.apps.filter(function (element) { return element.displayName == 'SharePoint'; });
                        OneNote = this.state.apps.filter(function (element) { return element.displayName == 'OneNote'; });
                        return [4 /*yield*/, this._installApp(teamId, SPAPP[0].id, graphClient)];
                    case 9:
                        isSPInstalled = _a.sent();
                        if (!isSPInstalled) {
                            this._onError();
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this._installApp(teamId, OneNote[0].id, graphClient)];
                    case 10:
                        isOneNoteInstalled = _a.sent();
                        if (!isOneNoteInstalled) {
                            this._onError();
                            return [2 /*return*/];
                        }
                        //
                        // タブ追加
                        //
                        this._addGroupTab(teamId, channelAId, 'SharePoint', SPAPP[0].id, graphClient);
                        this._addGroupTab(teamId, channelAId, 'OneNote', OneNote[0].id, graphClient);
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
                        return [3 /*break*/, 12];
                    case 11:
                        if (this.state.template == 'project') {
                        }
                        _a.label = 12;
                    case 12: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * 1.1 Creates O365 group
     * @param graphClient graph client
     */
    TeamCreator.prototype._createGroup = function (graphClient) {
        return __awaiter(this, void 0, void 0, function () {
            var displayName, mailNickname, _a, owners, members, groupRequest, response, error_1;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        displayName = this.state.teamName;
                        mailNickname = this._generateMailNickname(this.state.teamNickName);
                        _a = this.state, owners = _a.owners, members = _a.members;
                        groupRequest = {
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
                            groupRequest['owners@data.bind'] = owners.map(function (owner) {
                                return "https://graph.microsoft.com/v1.0/users/" + owner;
                            });
                        }
                        if (members && members.length) {
                            groupRequest['members@data.bind'] = members.map(function (member) {
                                return "https://graph.microsoft.com/v1.0/users/" + member;
                            });
                        }
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, graphClient.api('groups').version('v1.0').post(groupRequest)];
                    case 2:
                        response = _b.sent();
                        return [2 /*return*/, response.id];
                    case 3:
                        error_1 = _b.sent();
                        return [2 /*return*/, ''];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * 1.2 Creates team. as mentioned in the documentation - we need to make multiple attempts if team creation request errored
     * @param groupId group id
     * @param graphClient graph client
     */
    TeamCreator.prototype._createTeamWithAttempts = function (groupId, graphClient) {
        return __awaiter(this, void 0, void 0, function () {
            var attemptsCount, teamId;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        attemptsCount = 0;
                        teamId = '';
                        _a.label = 1;
                    case 1: return [4 /*yield*/, this._createTeam(groupId, graphClient)];
                    case 2:
                        teamId = _a.sent();
                        if (teamId) {
                            attemptsCount = 3;
                        }
                        else {
                            attemptsCount++;
                        }
                        _a.label = 3;
                    case 3:
                        if (attemptsCount < 3) return [3 /*break*/, 1];
                        _a.label = 4;
                    case 4: return [2 /*return*/, teamId];
                }
            });
        });
    };
    /**
     * Waits 10 seconds and tries to create a team
     * @param groupId group id
     * @param graphClient graph client
     */
    TeamCreator.prototype._createTeam = function (groupId, graphClient) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, new Promise(function (resolve) {
                        setTimeout(function () {
                            graphClient.api("groups/" + groupId + "/team").version('v1.0').put({
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
                            }).then(function (response) {
                                resolve(response.id);
                            }, function () {
                                resolve('');
                            });
                        }, 10000);
                    })];
            });
        });
    };
    /**
     * 1.3 Creates channel in the team
     * @param teamId team id
     * @param graphClient graph client
     */
    TeamCreator.prototype._createGroupChannel = function (teamId, channelName, channelDescription, graphClient) {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, graphClient.api("teams/" + teamId + "/channels").version('v1.0').post({
                                displayName: channelName,
                                description: channelDescription
                            })];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, response.id];
                    case 2:
                        error_2 = _a.sent();
                        console.error(error_2);
                        return [2 /*return*/, ''];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * 1.4 Installs the app to the team
     * @param teamId team Id
     * @param graphClient graph client
     */
    TeamCreator.prototype._installApp = function (teamId, selectedAppId, graphClient) {
        return __awaiter(this, void 0, void 0, function () {
            var error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, graphClient.api("teams/" + teamId + "/installedApps").version('v1.0').post({
                                'teamsApp@odata.bind': "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + selectedAppId
                            })];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        error_3 = _a.sent();
                        console.error(error_3);
                        return [2 /*return*/, false];
                    case 3: return [2 /*return*/, true];
                }
            });
        });
    };
    TeamCreator.prototype._addGroupTab = function (teamId, channelId, tabName, appId, graphClient) {
        return __awaiter(this, void 0, void 0, function () {
            var isTabCreated;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._addGroupTab2(teamId, channelId, tabName, appId, graphClient)];
                    case 1:
                        isTabCreated = _a.sent();
                        if (!isTabCreated) {
                            this._onError();
                        }
                        else {
                            this.setState({
                                creationState: CreationState.created
                            });
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    TeamCreator.prototype._onError = function (message) {
        this.setState({
            creationState: CreationState.error
        });
    };
    /**
     * 1.5 Adds tab to the specified channel of the team
     * @param teamId team id
     * @param channelId channel id
     * @param graphClient graph client
     */
    TeamCreator.prototype._addGroupTab2 = function (teamId, channelId, tabName, selectedAppId, graphClient) {
        return __awaiter(this, void 0, void 0, function () {
            var error_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, graphClient.api("teams/" + teamId + "/channels/" + channelId + "/tabs").version('v1.0').post({
                                displayName: tabName,
                                'teamsApp@odata.bind': "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + selectedAppId
                            })];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        error_4 = _a.sent();
                        console.error(error_4);
                        return [2 /*return*/, false];
                    case 3: return [2 /*return*/, true];
                }
            });
        });
    };
    /**
     * Generates mail nick name by display name of the group
     * @param displayName group display name
     */
    TeamCreator.prototype._generateMailNickname = function (displayName) {
        return displayName.toLowerCase().replace(/\s/gmi, '-');
    };
    return TeamCreator;
}(React.Component));
export default TeamCreator;
//# sourceMappingURL=TeamCreator.js.map