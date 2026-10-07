/**
 * Ambient module declarations for non-code imports handled by webpack loaders.
 *
 * `*.png` imports resolve to a file URL string (`asset/resource`).
 */

declare module "*.png" {
  const url: string;
  export default url;
}
