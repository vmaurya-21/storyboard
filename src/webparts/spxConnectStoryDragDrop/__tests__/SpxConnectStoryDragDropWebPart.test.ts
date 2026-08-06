import { Version } from '@microsoft/sp-core-library';
import SpxConnectStoryDragDropWebPart from '../SpxConnectStoryDragDropWebPart';

describe('SpxConnectStoryDragDropWebPart', () => {
  let webPart: SpxConnectStoryDragDropWebPart;

  beforeEach(() => {
    webPart = new SpxConnectStoryDragDropWebPart();
  });

  describe('Initialization', () => {
    it('should create instance', () => {
      expect(webPart).toBeDefined();
      expect(webPart).toBeInstanceOf(SpxConnectStoryDragDropWebPart);
    });

    it('should have dataVersion property', () => {
      const version = (webPart as any).dataVersion;
      expect(version).toBeDefined();
      expect(version).toBeInstanceOf(Version);
    });
  });

  describe('Property Pane Configuration', () => {
    it('should return property pane configuration', () => {
      const config = (webPart as any).getPropertyPaneConfiguration();
      expect(config).toBeDefined();
    });

    it('should have pages in configuration', () => {
      const config = (webPart as any).getPropertyPaneConfiguration();
      expect(config.pages).toBeDefined();
      expect(config.pages.length).toBeGreaterThan(0);
    });

    it('should have header in first page', () => {
      const config = (webPart as any).getPropertyPaneConfiguration();
      expect(config.pages[0].header).toBeDefined();
    });

    it('should have groups in first page', () => {
      const config = (webPart as any).getPropertyPaneConfiguration();
      expect(config.pages[0].groups).toBeDefined();
      expect(config.pages[0].groups.length).toBeGreaterThan(0);
    });

    it('should have description field', () => {
      const config = (webPart as any).getPropertyPaneConfiguration();
      const fields = config.pages[0].groups[0].groupFields;
      expect(fields).toBeDefined();
      expect(fields.length).toBeGreaterThan(0);
    });
  });

  describe('Properties', () => {
    it('should have default description property', () => {
      expect((webPart as any).properties).toBeDefined();
    });

    it('should allow setting description', () => {
      (webPart as any).properties.description = 'Test Description';
      expect((webPart as any).properties.description).toBe('Test Description');
    });
  });

  describe('Rendering', () => {
    it('should have render method', () => {
      expect(webPart.render).toBeDefined();
      expect(typeof webPart.render).toBe('function');
    });

    it('should have onDispose method', () => {
      expect((webPart as any).onDispose).toBeDefined();
      expect(typeof (webPart as any).onDispose).toBe('function');
    });
  });

  describe('Version', () => {
    it('should return correct data version', () => {
      const version = (webPart as any).dataVersion;
      expect(version.toString()).toBe('1.0');
    });
  });

  describe('Lifecycle Methods', () => {
    let domElement: HTMLDivElement;

    beforeEach(() => {
      domElement = document.createElement('div');
      document.body.appendChild(domElement);
      (webPart as any).domElement = domElement;

      (webPart as any).context = {
        pageContext: {
          user: {
            displayName: 'Test User'
          }
        },
        sdks: {
          microsoftTeams: null
        }
      };

      (webPart as any).properties = {
        description: 'Test Description'
      };
    });

    afterEach(() => {
      document.body.removeChild(domElement);
    });

    it('should call onInit successfully', async () => {
      const onInit = (webPart as any).onInit;
      const result = await onInit.call(webPart);

      expect(result).toBeUndefined();
    });

    it('should handle onThemeChanged with valid theme', () => {
      const mockTheme = {
        isInverted: true,
        semanticColors: {
          bodyText: '#ffffff',
          link: '#0078d4',
          linkHovered: '#106ebe'
        }
      };

      (webPart as any).onThemeChanged(mockTheme);

      expect(domElement.style.getPropertyValue('--bodyText')).toBe('#ffffff');
      expect(domElement.style.getPropertyValue('--link')).toBe('#0078d4');
      expect(domElement.style.getPropertyValue('--linkHovered')).toBe('#106ebe');
    });

    it('should handle onThemeChanged with undefined theme', () => {
      (webPart as any).onThemeChanged(undefined);

      expect(domElement.style.getPropertyValue('--bodyText')).toBe('');
    });

    it('should handle onThemeChanged with theme without semanticColors', () => {
      const mockTheme = {
        isInverted: false
      };

      (webPart as any).onThemeChanged(mockTheme);

      expect(domElement.style.getPropertyValue('--bodyText')).toBe('');
    });

    it('should call onDispose and unmount component', () => {
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      const unmountSpy = jest.spyOn(require('react-dom'), 'unmountComponentAtNode');

      (webPart as any).onDispose();

      expect(unmountSpy).toHaveBeenCalledWith(domElement);
      unmountSpy.mockRestore();
    });
  });
});
