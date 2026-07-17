// Mock @microsoft/sp-component-base
export interface IReadonlyTheme {
  isInverted?: boolean;
  semanticColors?: {
    bodyText?: string;
    link?: string;
    linkHovered?: string;
  };
}
