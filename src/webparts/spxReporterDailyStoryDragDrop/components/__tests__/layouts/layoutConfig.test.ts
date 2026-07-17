import { getLayoutConfig, LAYOUTS, AVAILABLE_LAYOUTS } from '../../layouts/layoutConfig';

describe('Layout Config', () => {
  describe('getLayoutConfig', () => {
    it('should return reporter daily layout config', () => {
      const config = getLayoutConfig('reporterDaily');
      
      expect(config.id).toBe('reporterDaily');
      expect(config.name).toBe('Reporter Daily');
      expect(config.slots).toHaveLength(5);
    });

    it('should return general layout config', () => {
      const config = getLayoutConfig('general');
      
      expect(config.id).toBe('general');
      expect(config.name).toBe('General');
      expect(config.slots).toHaveLength(6);
    });

    it('should return highlight layout config', () => {
      const config = getLayoutConfig('highlight');
      
      expect(config.id).toBe('highlight');
      expect(config.name).toBe('Highlight');
      expect(config.slots).toHaveLength(0);
    });
  });

  describe('Reporter Daily Layout', () => {
    const config = LAYOUTS.reporterDaily;

    it('should have correct slot IDs', () => {
      const slotIds = config.slots.map(slot => slot.id);
      
      expect(slotIds).toEqual([
        'rd-slot-1',
        'rd-slot-2',
        'rd-slot-3',
        'rd-slot-4',
        'rd-slot-5'
      ]);
    });

    it('should have correct variants', () => {
      expect(config.slots[0].variant).toBe('large');
      expect(config.slots[1].variant).toBe('horizontal');
      expect(config.slots[2].variant).toBe('small');
      expect(config.slots[3].variant).toBe('large');
      expect(config.slots[4].variant).toBe('fullHeight');
    });

    it('should have text-only slot without image', () => {
      const textOnlySlot = config.slots.find(slot => slot.showImage === false);
      
      expect(textOnlySlot).toBeDefined();
      expect(textOnlySlot?.id).toBe('rd-slot-3');
    });

    it('should have correct grid template', () => {
      expect(config.gridTemplate.columns).toBe('1fr 1fr 1fr');
      expect(config.gridTemplate.rows).toBe('60% 40%');
      expect(config.gridTemplate.areas).toBeDefined();
    });
  });

  describe('General Layout', () => {
    const config = LAYOUTS.general;

    it('should have 6 slots', () => {
      expect(config.slots).toHaveLength(6);
    });

    it('should have all general variant slots', () => {
      const allGeneral = config.slots.every(slot => slot.variant === 'general');
      
      expect(allGeneral).toBe(true);
    });

    it('should have correct slot IDs', () => {
      const slotIds = config.slots.map(slot => slot.id);
      
      expect(slotIds).toEqual([
        'gen-slot-1',
        'gen-slot-2',
        'gen-slot-3',
        'gen-slot-4',
        'gen-slot-5',
        'gen-slot-6'
      ]);
    });

    it('should have 2x3 grid template', () => {
      expect(config.gridTemplate.columns).toBe('1fr 1fr 1fr');
      expect(config.gridTemplate.rows).toBe('1fr 1fr');
    });
  });

  describe('AVAILABLE_LAYOUTS', () => {
    it('should include reporterDaily and general', () => {
      expect(AVAILABLE_LAYOUTS).toContain('reporterDaily');
      expect(AVAILABLE_LAYOUTS).toContain('general');
    });

    it('should not include highlight', () => {
      expect(AVAILABLE_LAYOUTS).not.toContain('highlight');
    });

    it('should have 2 available layouts', () => {
      expect(AVAILABLE_LAYOUTS).toHaveLength(2);
    });
  });
});
