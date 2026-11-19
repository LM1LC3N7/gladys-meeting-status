// -----------------------------------------------------------------------------
// Generates cover.png (800×534, the store size): an office door on a wall,
// a red "in a meeting" light glowing above it, a small headset on the right.
// No logo. Plain Node (zlib), no dependency.
// Run: npm run cover
// -----------------------------------------------------------------------------

import { writeFileSync } from 'node:fs';
import { deflateSync } from 'node:zlib';

const WIDTH = 800;
const HEIGHT = 534;

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k += 1) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  return c >>> 0;
});

function crc32(buffer) {
  let c = 0xffffffff;
  for (const byte of buffer) {
    c = CRC_TABLE[(c ^ byte) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([length, body, crc]);
}

const mix = (a, b, t) => a.map((v, i) => Math.round(v + (b[i] - v) * Math.min(1, Math.max(0, t))));
const smooth = (edge, distance) => Math.min(1, Math.max(0, edge - distance + 0.5));

const WALL_TOP = [44, 52, 70];
const WALL_BOTTOM = [70, 80, 102];
const FLOOR = [36, 38, 48];
const WOOD = [150, 104, 66];
const WOOD_DARK = [112, 74, 44];
const FRAME = [214, 216, 224];
const BOX = [28, 28, 34];
const RED = [235, 64, 70];
const GLOW = [255, 120, 110];
const KNOB = [230, 196, 92];
const HEADSET = [196, 200, 214];

/** Distance from a point to a rounded rectangle (negative inside). */
function box(x, y, cx, cy, hw, hh, r) {
  const qx = Math.abs(x - cx) - hw + r;
  const qy = Math.abs(y - cy) - hh + r;
  return Math.hypot(Math.max(qx, 0), Math.max(qy, 0)) + Math.min(Math.max(qx, qy), 0) - r;
}

function pixel(x, y) {
  let color = y < 440 ? mix(WALL_TOP, WALL_BOTTOM, y / 440) : FLOOR;
  // The glow of the light on the wall.
  const glow = Math.hypot(x - 330, (y - 92) * 1.6);
  color = mix(color, GLOW, Math.max(0, 0.55 - glow / 260));
  // The door frame, then the door and its panels.
  color = mix(color, FRAME, smooth(0, box(x, y, 330, 300, 118, 150, 4)));
  color = mix(color, WOOD, smooth(0, box(x, y, 330, 306, 104, 140, 2)));
  for (const py of [236, 366]) {
    color = mix(color, WOOD_DARK, smooth(0, box(x, y, 330, py, 74, 48, 6)) * 0.6);
  }
  color = mix(color, KNOB, smooth(0, Math.hypot(x - 410, y - 310) - 9));
  // The light box above the door, glowing red.
  color = mix(color, BOX, smooth(0, box(x, y, 330, 92, 92, 30, 10)));
  color = mix(color, RED, smooth(0, box(x, y, 330, 92, 80, 20, 8)));
  color = mix(
    color,
    GLOW,
    Math.max(0, 0.7 - Math.abs(y - 86) / 14) * smooth(0, box(x, y, 330, 92, 70, 14, 6)),
  );
  // A headset on the right: the band (a half ring) and two ear cups.
  const ring = Math.abs(Math.hypot(x - 610, y - 300) - 80);
  if (y < 300) {
    color = mix(color, HEADSET, smooth(0, ring - 9));
  }
  for (const ex of [530, 690]) {
    color = mix(color, HEADSET, smooth(0, box(x, y, ex, 318, 22, 38, 14)));
    color = mix(color, BOX, smooth(0, box(x, y, ex, 318, 12, 26, 8)) * 0.8);
  }
  // A dot on the headset: in a meeting.
  color = mix(color, RED, smooth(0, Math.hypot(x - 690, y - 250) - 12));
  return color;
}

const raw = Buffer.alloc((WIDTH * 3 + 1) * HEIGHT);
for (let y = 0; y < HEIGHT; y += 1) {
  const row = y * (WIDTH * 3 + 1);
  raw[row] = 0;
  for (let x = 0; x < WIDTH; x += 1) {
    const [r, g, b] = pixel(x, y);
    raw.set([r, g, b], row + 1 + x * 3);
  }
}

const header = Buffer.alloc(13);
header.writeUInt32BE(WIDTH, 0);
header.writeUInt32BE(HEIGHT, 4);
header[8] = 8; // bit depth
header[9] = 2; // RGB
const png = Buffer.concat([
  Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
  chunk('IHDR', header),
  chunk('IDAT', deflateSync(raw, { level: 9 })),
  chunk('IEND', Buffer.alloc(0)),
]);
writeFileSync(new URL('../cover.png', import.meta.url), png);
console.log(`cover.png: ${WIDTH}×${HEIGHT}, ${png.length} bytes`);
