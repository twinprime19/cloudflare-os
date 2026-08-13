// `crypto.subtle.timingSafeEqual` is a workerd extension that Node's WebCrypto does not implement.
// The nonce comparison in the gatekeeper uses it, so supply an equivalent (non-constant-time, which
// is irrelevant to a test) rather than reshaping production code around the test runtime.

type TimingSafeEqual = (a: ArrayBufferView, b: ArrayBufferView) => boolean;

const subtle = crypto.subtle as SubtleCrypto & { timingSafeEqual?: TimingSafeEqual };

if (typeof subtle.timingSafeEqual !== "function") {
  Object.defineProperty(subtle, "timingSafeEqual", {
    configurable: true,
    writable: true,
    value: (a: ArrayBufferView, b: ArrayBufferView): boolean => {
      const left = new Uint8Array(a.buffer, a.byteOffset, a.byteLength);
      const right = new Uint8Array(b.buffer, b.byteOffset, b.byteLength);
      if (left.length !== right.length) return false;
      return left.every((byte, index) => byte === right[index]);
    },
  });
}
