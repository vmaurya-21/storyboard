// Mock @microsoft/sp-webpart-base
import { Version } from '@microsoft/sp-core-library';

export const mockContext = {
  pageContext: {
    web: {
      absoluteUrl: 'https://contoso.sharepoint.com',
      title: 'Test Site',
      serverRelativeUrl: '/sites/test'
    },
    user: {
      displayName: 'Test User',
      email: 'test@contoso.com',
      loginName: 'i:0#.f|membership|test@contoso.com'
    },
    list: undefined,
    listItem: undefined
  },
  sdks: {
    microsoftTeams: undefined
  },
  spHttpClient: {
    get: jest.fn(),
    post: jest.fn(),
    fetch: jest.fn()
  }
};

export class BaseClientSideWebPart {
  public context = mockContext;
  public properties: any = {};
  public domElement: HTMLElement = document.createElement('div');
  private _dataVersion = Version.parse('1.0');
  
  public render(): void {}
  protected onInit(): Promise<void> { return Promise.resolve(); }
  protected onDispose(): void {}
  protected get dataVersion() { return this._dataVersion; }
}

export const PropertyPaneTextField = jest.fn();

export interface IPropertyPaneConfiguration {
  pages: any[];
}
