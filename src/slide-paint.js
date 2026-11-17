// DrawingML angles are clockwise from the positive X axis. CSS starts at up.
export function gradientCss(paint, width, height) {
  if (paint?.type !== 'linear' || !paint.stops?.length) return '';
  let angle = Number(paint.angle) || 0;
  if (paint.scaled && width > 0 && height > 0) {
    const radians = angle * Math.PI / 180;
    angle = Math.atan2(height * Math.sin(radians), width * Math.cos(radians)) * 180 / Math.PI;
  }
  const stops = paint.stops.map(stop => `${colorOpacity(stop.color, stop.opacity)} ${Math.max(0, Math.min(1, stop.offset)) * 100}%`);
  // Blend in linear light rather than mixing encoded channel values.
  return `linear-gradient(${angle + 90}deg in srgb-linear, ${stops.join(', ')})`;
}

export function colorOpacity(color, opacity) {
  if (opacity === undefined || opacity >= 1 || !/^#[0-9a-f]{6}$/i.test(color || '')) return color;
  const rgb = [1, 3, 5].map(offset => parseInt(color.slice(offset, offset + 2), 16));
  return `rgba(${rgb.join(',')},${Math.max(0, opacity)})`;
}
