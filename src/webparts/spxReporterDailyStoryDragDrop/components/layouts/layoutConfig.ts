import { LayoutConfig, LayoutType } from '../types';

// Reporter Daily Layout — ported 1:1 from the reference:
//   Column 1 (two-medium): full-image medium + thumbnail-text medium
//   Column 2 (small-tall):  text-only small   + full-image tall
//   Column 3 (full):        full-image large (fills the column)
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

// General Layout - 6 equal cards in 2 rows
// Row 1: 3 equal cards
// Row 2: 3 equal cards
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

// Highlight Layout - 5-slot layout with 4 medium cards and 1 tall card
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

// Connect Homepage Layout - a 5-slot full-screen carousel (original Highlight layout)
const connectHomepageLayout: LayoutConfig = {
  id: 'connectHomepage',
  name: 'Connect Homepage',
  description: 'Full-width carousel of up to 5 featured stories',
  slots: [
    { id: 'slot-1', variant: 'large', cardLayout: 'full-image' },
    { id: 'slot-2', variant: 'large', cardLayout: 'full-image' },
    { id: 'slot-3', variant: 'large', cardLayout: 'full-image' },
    { id: 'slot-4', variant: 'large', cardLayout: 'full-image' },
    { id: 'slot-5', variant: 'large', cardLayout: 'full-image' }
  ],
  gridTemplate: {
    columns: '1fr',
    rows: '1fr',
    areas: ''
  }
};

// Layout Registry
export const LAYOUTS: Record<LayoutType, LayoutConfig> = {
  reporterDaily: reporterDailyLayout,
  general: generalLayout,
  highlight: highlightLayout,
  connectHomepage: connectHomepageLayout
};

// Helper function to get layout by type
export const getLayoutConfig = (layoutType: LayoutType): LayoutConfig => {
  return LAYOUTS[layoutType];
};

// Available layouts for the "View" dropdown — the layouts.
export const AVAILABLE_LAYOUTS: LayoutType[] = ['reporterDaily', 'general', 'highlight', 'connectHomepage'];
