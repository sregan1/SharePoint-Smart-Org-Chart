import { WebPartContext } from '@microsoft/sp-webpart-base';

/** Replaces `{token}`-style placeholders in a loc string with the given values. */
export function formatString(template: string, params: { [key: string]: string | number }): string {
  return template.replace(/\{(\w+)\}/g, (match, key) => {
    const value = params[key];
    return value === undefined ? match : String(value);
  });
}

/** The SharePoint page's UI culture, for locale-aware date/number formatting. */
export function getEffectiveLocale(context: WebPartContext | undefined): string {
  return context?.pageContext?.cultureInfo?.currentUICultureName || 'en-US';
}
