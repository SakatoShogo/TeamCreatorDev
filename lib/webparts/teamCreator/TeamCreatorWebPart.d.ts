import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
export interface ITeamCreatorWebPartProps {
    description: string;
}
export default class TeamCreatorWebPart extends BaseClientSideWebPart<ITeamCreatorWebPartProps> {
    render(): void;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=TeamCreatorWebPart.d.ts.map