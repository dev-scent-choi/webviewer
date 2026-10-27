import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';

const SAFE_COLOR = /^#[0-9a-f]{6}$/i;

/**
 * 연결이 보낸 접속 정보(awareness) 갱신을 걸러 다시 만든다 — 서버가 그대로 퍼뜨리면 열람만 하는 연결도
 * 다른 참여자를 지우거나(그 사람 clientID로 null) 이름을 사칭하고, 커서 색 자리에 CSS를 넣어
 * 다른 사람 화면을 덮을 수 있었다(y-prosemirror 기본 커서가 색을 style 속성에 그대로 넣는다).
 * - 다른 연결이 쓰는 clientID 의 항목은 버린다
 * - user.name 은 토큰에 서명된 이름으로 덮고, user.color 는 #rrggbb 만 남긴다
 * 형식이 y-protocols encodeAwarenessUpdate 와 같다: 개수, (clientID, clock, JSON 상태)…
 */
export function sanitizeAwarenessUpdate(update: Uint8Array, foreignIds: ReadonlySet<number>, viewer: string | null): Uint8Array {
  const decoder = decoding.createDecoder(update);
  const entries: { id: number; clock: number; state: string }[] = [];
  const count = decoding.readVarUint(decoder);
  for (let i = 0; i < count; i += 1) {
    const id = decoding.readVarUint(decoder);
    const clock = decoding.readVarUint(decoder);
    const raw = decoding.readVarString(decoder);
    if (foreignIds.has(id)) continue;
    const state: unknown = JSON.parse(raw);
    if (state && typeof state === 'object' && 'user' in state) {
      const user = (state as { user: unknown }).user;
      const color = user && typeof user === 'object' ? (user as { color?: unknown }).color : undefined;
      const name = viewer ?? (user && typeof user === 'object' ? String((user as { name?: unknown }).name ?? '') : '');
      (state as { user: unknown }).user = { name: name.slice(0, 80), ...(typeof color === 'string' && SAFE_COLOR.test(color) ? { color } : {}) };
    }
    entries.push({ id, clock, state: JSON.stringify(state) });
  }
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, entries.length);
  for (const e of entries) {
    encoding.writeVarUint(encoder, e.id);
    encoding.writeVarUint(encoder, e.clock);
    encoding.writeVarString(encoder, e.state);
  }
  return encoding.toUint8Array(encoder);
}
