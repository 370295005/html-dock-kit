// Juice ships declarations for its Node entry but not the client subpath.
// Declare only the synchronous API used here; resource-fetching APIs are not
// present on juice/client and are deliberately excluded from this declaration.
declare module 'juice/client' {
  import type { Options } from 'juice';
  import type { CheerioAPI } from 'cheerio/slim';

  interface JuiceClient {
    (html: string, options?: Options): string;
    inlineContent(html: string, css: string, options?: Options): string;
    inlineDocument(document: CheerioAPI, css: string, options?: Options): Set<string> | null;
  }
  const juiceClient: JuiceClient;
  export default juiceClient;
}
