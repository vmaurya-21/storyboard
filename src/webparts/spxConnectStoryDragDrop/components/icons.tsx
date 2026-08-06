import * as React from 'react';

/**
 * Props shared by every icon in this module.
 */
export interface IIconProps {
  /** Square edge length in pixels. Each icon defines its own default. */
  size?: number;
  /** Class applied to the `<svg>` element. */
  className?: string;
}

/**
 * Builds the SVG attributes every icon shares.
 *
 * The values match Lucide's own defaults — a 24x24 view box, no fill, a
 * 2px `currentColor` stroke with round caps and joins — so an icon inherits
 * its colour from the surrounding text. Icons are decorative, so they are
 * hidden from assistive technology and removed from the tab order; the
 * surrounding control supplies the accessible name.
 *
 * @param size - Rendered width and height in pixels.
 * @returns Attributes to spread onto an `<svg>` element.
 */
const base = (size: number): React.SVGProps<SVGSVGElement> => ({
  xmlns: 'http://www.w3.org/2000/svg',
  width: size,
  height: size,
  viewBox: '0 0 24 24',
  fill: 'none',
  stroke: 'currentColor',
  strokeWidth: 2,
  strokeLinecap: 'round' as const,
  strokeLinejoin: 'round' as const,
  'aria-hidden': true,
  focusable: false
});

/**
 * LinkedIn glyph, shown on the source badge of a story linking to LinkedIn.
 *
 * Path data is Lucide `linkedin` (lucide-react 0.487.0), inlined because the
 * Fluent icon set has no faithful equivalent. Defaults to 10px to match the
 * badge's `w-2.5` in the reference.
 */
export const LinkedinIcon: React.FC<IIconProps> = ({ size = 10, className }) => (
  <svg {...base(size)} className={className}>
    <path d="M16 8a6 6 0 0 1 6 6v7h-4v-7a2 2 0 0 0-2-2 2 2 0 0 0-2 2v7h-4v-7a6 6 0 0 1 6-6z" />
    <rect width="4" height="12" x="2" y="9" />
    <circle cx="4" cy="4" r="2" />
  </svg>
);

/**
 * Outbound-link glyph, shown on the source badge of a story pointing off-tenant.
 *
 * Path data is Lucide `external-link` (lucide-react 0.487.0). Defaults to 10px
 * to match the badge's `w-2.5` in the reference.
 */
export const ExternalLinkIcon: React.FC<IIconProps> = ({ size = 10, className }) => (
  <svg {...base(size)} className={className}>
    <path d="M15 3h6v6" />
    <path d="M10 14 21 3" />
    <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" />
  </svg>
);

/**
 * Close glyph for dialog headers.
 *
 * Path data is Lucide `x` (lucide-react 0.487.0). Defaults to 16px to match
 * the `size-4` the reference's dialog close control renders at.
 */
export const XIcon: React.FC<IIconProps> = ({ size = 16, className }) => (
  <svg {...base(size)} className={className}>
    <path d="M18 6 6 18" />
    <path d="m6 6 12 12" />
  </svg>
);

/**
 * Pencil-on-square glyph for the edit action on an Available Stories row.
 *
 * Path data is Lucide `square-pen`, which `Edit` aliases in lucide-react
 * 0.487.0. Defaults to 12px to match the reference's `w-3`.
 */
export const EditIcon: React.FC<IIconProps> = ({ size = 12, className }) => (
  <svg {...base(size)} className={className}>
    <path d="M12 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
    <path d="M18.375 2.625a2.121 2.121 0 1 1 3 3L12 15l-4 1 1-4Z" />
  </svg>
);
