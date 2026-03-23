declare module 'utif' {
  export function decode(buffer: ArrayBuffer): any[];
  export function decodeImage(img: any, ifds: any[]): void;
  export function toRGBA8(img: any): Uint8Array;
}
