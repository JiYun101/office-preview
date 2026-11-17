export const utf16 = new TextDecoder('utf-16le');

export function binary(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const need = (offset, length) => {
    if (!Number.isSafeInteger(offset) || offset < 0 || length < 0 || offset + length > bytes.length) {
      throw new Error('旧版 Office 文件中的记录长度或偏移越界，文件可能已损坏。');
    }
  };
  return {
    bytes, need,
    u8(offset) { need(offset, 1); return view.getUint8(offset); },
    u16(offset) { need(offset, 2); return view.getUint16(offset, true); },
    u32(offset) { need(offset, 4); return view.getUint32(offset, true); },
    f64(offset) { need(offset, 8); return view.getFloat64(offset, true); },
    slice(offset, length) { need(offset, length); return bytes.subarray(offset, offset + length); }
  };
}

export function warn(warnings, message) {
  if (!warnings.includes(message)) warnings.push(message);
}

export function lowUnicode(bytes) {
  // In BIFF/PPT compressed Unicode each byte is a UTF-16 low byte, not ANSI.
  let result = '';
  for (let start = 0; start < bytes.length; start += 8192) {
    result += String.fromCharCode(...bytes.subarray(start, start + 8192));
  }
  return result;
}

