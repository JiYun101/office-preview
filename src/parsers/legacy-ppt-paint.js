import { binary, warn } from './legacy-binary.js';

const property = (properties, id, fallback) => properties.get(id)?.value ?? fallback;
const clamp = value => Math.max(0, Math.min(1, value));

// OfficeArt angles rotate counterclockwise from its bottom-to-top vector.
// Our normalized ramp is foreground-to-background, with DrawingML clockwise
// angles from left-to-right. The focus maps the background end of that ramp.
export function readLegacyFillPaint(properties, resolveColor, warnings) {
  const type = property(properties, 0x0180, 0);
  if (![4, 7].includes(type)) return undefined;
  const front = resolveColor(property(properties, 0x0181, 0x00ffffff));
  const back = resolveColor(property(properties, 0x0183, 0x00ffffff));
  const frontOpacity = clamp(property(properties, 0x0182, 65536) / 65536);
  const backOpacity = clamp(property(properties, 0x0184, 65536) / 65536);
  let stops = [{ offset: 0, color: front, opacity: frontOpacity }, { offset: 1, color: back, opacity: backOpacity }];
  const shade = properties.get(0x0197);
  if (shade?.complex) {
    const data = binary(shade.bytes);
    const count = data.u16(0), allocated = data.u16(2), size = data.u16(4);
    if (size !== 8 || count > allocated || count > 4096 || shade.bytes.length !== 6 + count * 8) throw new Error('PPT 渐变颜色数组长度无效。');
    if (count) {
      stops = [];
      let last = -1;
      for (let i = 0; i < count; i++) {
        const offset = data.u32(10 + i * 8) / 65536;
        if (offset < last || offset < 0 || offset > 1) throw new Error('PPT 渐变颜色位置无效。');
        stops.push({ offset, color: resolveColor(data.u32(6 + i * 8)), opacity: frontOpacity + (backOpacity - frontOpacity) * offset });
        last = offset;
      }
      if (stops[0].offset !== 0) stops.unshift({ ...stops[0], offset: 0 });
      if (stops.at(-1).offset !== 1) stops.push({ ...stops.at(-1), offset: 1 });
    }
  }
  const angle = ((90 - (property(properties, 0x018b, 0) | 0) / 65536) % 360 + 360) % 360;
  let focus = property(properties, 0x018c, 0) | 0;
  if (focus < -100 || focus > 100) throw new Error('PPT 渐变焦点超出有效范围。');
  if (focus < 0) {
    // Negative focal points require a different reflected-color envelope.
    warn(warnings, '部分 PPT 负焦点渐变当前按基本反射渐变近似显示。');
    focus = -focus;
  }
  if (focus === 0) stops = [...stops].reverse().map(stop => ({ ...stop, offset: 1 - stop.offset }));
  else if (focus !== 100) {
    const center = focus / 100;
    stops = [...stops.map(stop => ({ ...stop, offset: stop.offset * center })), ...[...stops].reverse().slice(1).map(stop => ({ ...stop, offset: center + (1 - stop.offset) * (1 - center) }))];
  }
  return { type: 'linear', angle, scaled: type === 7, stops };
}

export function legacyCornerRadius(properties, width, height) {
  const adjustment = Math.max(0, Math.min(10800, property(properties, 0x0147, 3600)));
  return Math.max(0, Math.min(width, height)) * adjustment / 21600;
}
