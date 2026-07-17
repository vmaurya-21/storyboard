export interface IPropertyPaneConfiguration {
  pages: IPropertyPanePage[];
}

export interface IPropertyPanePage {
  header?: {
    description?: string;
  };
  groups: IPropertyPaneGroup[];
}

export interface IPropertyPaneGroup {
  groupName?: string;
  groupFields: any[];
}

export function PropertyPaneTextField(targetProperty: string, properties: any): any {
  return {
    type: 'TextField',
    targetProperty,
    properties
  };
}
