/** All lengths are in twips: 1 inch = 1440 twips; 1 point = 20 twips. */
export interface Margins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
  gutter: number;
}

/** Supply portrait dimensions; orientation controls their ordering. */
export type PageSize = 'letter' | 'A4' | { width: number; height: number };

export interface DocumentProperties {
  title?: string;
  creator?: string;
  description?: string;
}

export interface HtmlDocxOptions {
  orientation?: 'portrait' | 'landscape';
  pageSize?: PageSize;
  margins?: Partial<Margins>;
  /** External images are never fetched. Defaults to 'error'. */
  externalImages?: 'error' | 'preserve';
  /** Inline static <style> rules before Word import. Defaults to true. */
  inlineStyles?: boolean;
  /** ZIP compression level, integer 0-9. Defaults to 6. */
  compression?: number;
  properties?: DocumentProperties;
}
