/** The archive's inflate module with the TypeScript inflater as its default (see `own-inflate.ts`). */
export * from "../../src/internal/zip/inflate.ts"
export { inflateRawSync as inflateRawWeb } from "../../src/internal/zip/rfc1951.ts"
