import '@webcomponents/custom-elements/src/native-shim';
import '@webcomponents/webcomponentsjs/bundles/webcomponents-sd-ce';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import '../../../../ang-spfx-rt-proj/dist/ang-spfx-rt-proj/main.NU6CW42L.js';
import '../../../../ang-spfx-rt-proj/dist/ang-spfx-rt-proj/polyfills.NDQMOGQV.js';
import '../../../../ang-spfx-rt-proj/dist/ang-spfx-rt-proj/styles.VFXLKGBH.js';
export interface ISpfxAngWebpartWebPartProps {
    description: string;
}
export default class SpfxAngWebpartWebPart extends BaseClientSideWebPart<ISpfxAngWebpartWebPartProps> {
    private _isDarkTheme;
    private _environmentMessage;
    render(): void;
    protected onInit(): Promise<void>;
    private _getEnvironmentMessage;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=SpfxAngWebpartWebPart.d.ts.map