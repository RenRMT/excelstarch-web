/**
 * Ambient module declarations for non-code imports handled by webpack loaders.
 *
 * `*.png?inline` is emitted by webpack as a base64 data-URI string (`asset/inline`); the chrome
 * logo imports it that way to feed Office.js `shape.addImage`. Plain `*.png` imports resolve to a
 * file URL string (`asset/resource`).
 */
declare module "*.png?inline" {
  const dataUri: string;
  export default dataUri;
}

declare module "*.png" {
  const url: string;
  export default url;
}
