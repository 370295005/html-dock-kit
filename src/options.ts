import type { HtmlDocxOptions, Margins } from './types.js';

export interface ResolvedOptions {
  width: number;
  height: number;
  orientation: 'portrait' | 'landscape';
  margins: Margins;
  externalImages: 'error' | 'preserve';
  inlineStyles: boolean;
  compression: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
  properties: { title: string; creator: string; description: string };
}

const DEFAULT_MARGINS: Margins = {
  top: 1440, right: 1440, bottom: 1440, left: 1440,
  header: 720, footer: 720, gutter: 0,
};

function object(value: unknown, name: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError(`${name} must be an object.`);
  }
  return value as Record<string, unknown>;
}

function length(value: unknown, name: string, allowZero: boolean): number {
  if (typeof value !== 'number' || !Number.isInteger(value) ||
      value < (allowZero ? 0 : 1) || value > 2147483647) {
    throw new RangeError(`${name} must be a ${allowZero ? 'non-negative' : 'positive'} integer in twips (at most 2147483647).`);
  }
  return value;
}

export function resolveOptions(input: HtmlDocxOptions): ResolvedOptions {
  const options = object(input, 'options');
  const orientation = options.orientation === undefined ? 'portrait' : options.orientation;
  if (orientation !== 'portrait' && orientation !== 'landscape') {
    throw new TypeError('orientation must be "portrait" or "landscape".');
  }
  const size = options.pageSize === undefined ? 'letter' : options.pageSize;
  let dimensions: { width: number; height: number };
  if (size === 'letter') dimensions = { width: 12240, height: 15840 };
  else if (size === 'A4') dimensions = { width: 11906, height: 16838 };
  else {
    const custom = object(size, 'pageSize');
    dimensions = {
      width: length(custom.width, 'pageSize.width', false),
      height: length(custom.height, 'pageSize.height', false),
    };
  }
  const shorter = Math.min(dimensions.width, dimensions.height);
  const longer = Math.max(dimensions.width, dimensions.height);
  const width = orientation === 'portrait' ? shorter : longer;
  const height = orientation === 'portrait' ? longer : shorter;
  const margins = { ...DEFAULT_MARGINS };
  if (options.margins !== undefined) {
    const values = object(options.margins, 'margins');
    for (const key of Object.keys(values)) {
      if (!Object.prototype.hasOwnProperty.call(DEFAULT_MARGINS, key)) throw new TypeError(`Unknown margin: ${key}.`);
      if (values[key] !== undefined) margins[key as keyof Margins] = length(values[key], `margins.${key}`, true);
    }
  }
  if (margins.left + margins.right + margins.gutter >= width || margins.top + margins.bottom >= height) {
    throw new RangeError('Page margins must leave a positive content area.');
  }
  const externalImages = options.externalImages === undefined ? 'error' : options.externalImages;
  if (externalImages !== 'error' && externalImages !== 'preserve') {
    throw new TypeError('externalImages must be "error" or "preserve".');
  }
  const inlineStyles = options.inlineStyles === undefined ? true : options.inlineStyles;
  if (typeof inlineStyles !== 'boolean') {
    throw new TypeError('inlineStyles must be a boolean.');
  }
  const compression = options.compression === undefined ? 6 : options.compression;
  if (typeof compression !== 'number' || !Number.isInteger(compression) || compression < 0 || compression > 9) {
    throw new RangeError('compression must be an integer from 0 to 9.');
  }
  const properties = { title: '', creator: '', description: '' };
  if (options.properties !== undefined) {
    const values = object(options.properties, 'properties');
    for (const key of Object.keys(properties) as (keyof typeof properties)[]) {
      const value = values[key];
      if (value !== undefined) {
        if (typeof value !== 'string') throw new TypeError(`properties.${key} must be a string.`);
        properties[key] = value;
      }
    }
  }
  return { width, height, orientation, margins, externalImages, inlineStyles,
    compression: compression as ResolvedOptions['compression'], properties };
}
