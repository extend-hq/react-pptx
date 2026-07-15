declare module '../pkg/pptx_wasm.js' {
  export default function init(source?: unknown): Promise<unknown>;
  export function parse_presentation(bytes: Uint8Array): unknown;
  export function detect_presentation_format(bytes: Uint8Array): string;
}
