import { LayoutConfig, LayoutType } from '../types';

/**
 * Reporter Daily preset: five slots across three columns.
 *
 * Columns one and two stack two slots each; column three holds a single
 * full-height slot spanning both rows.
 */
const reporterDailyLayout: LayoutConfig = {
  id: 'reporterDaily',
  name: 'Reporter Daily Site',
  description: 'Mixed layout with different slot sizes',
  slots: [
    { id: 'slot-1', variant: 'medium', cardLayout: 'full-image', gridArea: 'col1-top' },
    { id: 'slot-2', variant: 'medium', cardLayout: 'thumbnail-text', gridArea: 'col1-bottom' },
    { id: 'slot-3', variant: 'small', cardLayout: 'text-only', showImage: false, gridArea: 'col2-top' },
    { id: 'slot-4', variant: 'tall', cardLayout: 'full-image', gridArea: 'col2-bottom' },
    { id: 'slot-5', variant: 'large', cardLayout: 'full-image', gridArea: 'col3-full' }
  ],
  gridTemplate: {
    columns: '1fr 1fr 1fr',
    rows: 'auto auto',
    areas: `
      "col1-top col2-top col3-full"
      "col1-bottom col2-bottom col3-full"
    `
  }
};

/** General preset: six equal cards in a 3-column, 2-row grid. */
const generalLayout: LayoutConfig = {
  id: 'general',
  name: 'General',
  description: '6 equal-sized cards in a 2x3 grid',
  slots: [
    { id: 'gen-slot-1', variant: 'general', gridArea: 'card1' },
    { id: 'gen-slot-2', variant: 'general', gridArea: 'card2' },
    { id: 'gen-slot-3', variant: 'general', gridArea: 'card3' },
    { id: 'gen-slot-4', variant: 'general', gridArea: 'card4' },
    { id: 'gen-slot-5', variant: 'general', gridArea: 'card5' },
    { id: 'gen-slot-6', variant: 'general', gridArea: 'card6' }
  ],
  gridTemplate: {
    columns: '1fr 1fr 1fr',
    rows: '1fr 1fr',
    areas: `
      "card1 card2 card3"
      "card4 card5 card6"
    `
  }
};

/**
 * Highlight preset: five slots, one of them full height.
 *
 * The grid template is intentionally minimal — this preset is laid out by
 * {@link LayoutRenderer} rather than by named grid areas.
 */
const highlightLayout: LayoutConfig = {
  id: 'highlight',
  name: 'Highlight',
  description: '5 stories layout with one full-height slot',
  slots: [
    { id: 'slot-1', variant: 'medium', cardLayout: 'full-image' },
    { id: 'slot-2', variant: 'medium', cardLayout: 'full-image' },
    { id: 'slot-3', variant: 'large', cardLayout: 'full-image' },
    { id: 'slot-4', variant: 'medium', cardLayout: 'full-image' },
    { id: 'slot-5', variant: 'medium', cardLayout: 'full-image' }
  ],
  gridTemplate: {
    columns: '1fr',
    rows: '1fr',
    areas: ''
  }
};

/**
 * Connect Homepage preset: the five slots are presented as one full-width
 * carousel rather than a grid.
 *
 * Only the slot *ids* are read — `LayoutRenderer` hands them to `StoryCarousel`
 * and nothing else here reaches the DOM. The carousel sizes its own slides, so
 * `variant` is inert (it is required by `CardSlotConfig` and set to `large`
 * only because that is the closest description of a full-bleed slide), and the
 * grid template goes unused.
 *
 * Note there is no `cardLayout`: per-slide layout is an editor choice held in
 * `slotLayoutPreferences` and persisted with the board, so a default declared
 * here would be silently ignored.
 */
const connectHomepageLayout: LayoutConfig = {
  id: 'connectHomepage',
  name: 'Connect Homepage',
  description: 'Full-width carousel of up to 5 stories',
  slots: [
    { id: 'slot-1', variant: 'large' },
    { id: 'slot-2', variant: 'large' },
    { id: 'slot-3', variant: 'large' },
    { id: 'slot-4', variant: 'large' },
    { id: 'slot-5', variant: 'large' }
  ],
  gridTemplate: {
    columns: '1fr',
    rows: '1fr',
    areas: ''
  }
};

/** Every layout preset, keyed by {@link LayoutType}. */
export const LAYOUTS: Record<LayoutType, LayoutConfig> = {
  reporterDaily: reporterDailyLayout,
  general: generalLayout,
  highlight: highlightLayout,
  connectHomepage: connectHomepageLayout
};

/**
 * Looks up a layout preset by id.
 *
 * @param layoutType - Preset to resolve.
 * @returns The matching preset. Because {@link LAYOUTS} is a total record over
 * {@link LayoutType}, every valid id resolves.
 */
export const getLayoutConfig = (layoutType: LayoutType): LayoutConfig => {
  return LAYOUTS[layoutType];
};

/**
 * Presets offered in the layout picker, in display order.
 *
 * A subset of {@link LAYOUTS}: the others remain defined so existing boards
 * and scheduled groups saved against them still render, but they are no
 * longer selectable.
 */
export const AVAILABLE_LAYOUTS: LayoutType[] = ['connectHomepage'];
