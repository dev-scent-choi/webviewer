import * as Y from 'yjs';
import type { WebsocketProvider } from 'y-websocket';
import * as decoding from 'lib0/decoding';
import type * as encoding from 'lib0/encoding';
import { MESSAGE_ROOM_EPOCH, MESSAGE_ROOM_STALE, ROOM_EPOCH_PARAM } from '../collab/roomEpoch.js';

// eslint-disable-next-line @typescript-eslint/no-explicit-any -- Y 타입들의 공통 부모
type AnyType = Y.AbstractType<any>;
type Handler = (encoder: encoding.Encoder, decoder: decoding.Decoder) => void;

/** 새로 고침을 저절로 하기 전 보여 주는 초읽기(초) — 잃을 입력이 없을 때만 */
const AUTO_RELOAD_SECONDS = 10;
const MAX_RECOVERY_CHARS = 200_000;

/**
 * 편집기(워드·엑셀·PPT·PDF) 공용 — 방 세대 확인(src/collab/roomEpoch.ts 설명). 서버가 알려 준 세대를 다음 접속 주소에 싣고,
 * 이전 세대로 다시 접속했다고 거절당하면 다시 접속하지 않고(옛 상태를 새 방에 합치지 않는다) 안내한다.
 * 그 편집은 새 방의 문서와 안전하게 맞출 방법이 없어 다시 보내지 않는다(옛 세대의 Yjs 변경을 새 세대 문서에 다시
 * 적용하면 중복이 생긴다). 대신 이 화면에서 고친 곳(문단·셀)의 현재 글자를 모아 복사·내려받기로 건넨다.
 * 반쯤 끊긴 연결은 보낸 변경이 서버에 닿았는지 알 수 없으므로, 고친 기록은 접속 상태와 상관없이 모두 남긴다.
 */
export function guardRoomEpoch(provider: WebsocketProvider, ydoc: Y.Doc): void {
  const edits = new Map<AnyType, Set<string | null>>();
  ydoc.on('afterTransaction', (tr: Y.Transaction) => {
    if (tr.origin === provider) return; // 서버에서 온 변경
    tr.changed.forEach((keys, type) => {
      const seen = edits.get(type) ?? new Set<string | null>();
      keys.forEach((k) => seen.add(k));
      edits.set(type, seen);
    });
  });
  const handlers = provider.messageHandlers as unknown as Handler[];
  handlers[MESSAGE_ROOM_EPOCH] = (_encoder, decoder) => {
    provider.params[ROOM_EPOCH_PARAM] = decoding.readVarString(decoder);
  };
  handlers[MESSAGE_ROOM_STALE] = () => {
    provider.shouldConnect = false;
    provider.disconnect();
    showStaleRoomNotice(recoveryText(edits));
  };
}

/** 고친 Y 타입들의 현재 글자 — 문단은 글자만, 맵 칸은 '이름: 값' 한 줄씩 */
export function recoveryText(edits: Map<AnyType, Set<string | null>>): string {
  const lines: string[] = [];
  const seenNodes = new Set<AnyType>();
  for (const [type, keys] of edits) {
    if (type instanceof Y.Map) {
      for (const key of keys) {
        if (key === null) continue;
        const value = type.get(key);
        if (value === undefined) continue;
        lines.push(`${key}: ${typeof value === 'string' ? value : value instanceof Y.AbstractType ? textOf(value) : JSON.stringify(value)}`);
      }
      continue;
    }
    // 문단 단위로 모은다(글자 조각이 바뀌면 그 문단 전체)
    const node = type instanceof Y.XmlText && type.parent instanceof Y.XmlElement ? type.parent : type;
    if (seenNodes.has(node) || node.doc === null) continue;
    seenNodes.add(node);
    const text = textOf(node).trim();
    if (text) lines.push(text);
  }
  const all = lines.join('\n');
  return all.length > MAX_RECOVERY_CHARS ? all.slice(0, MAX_RECOVERY_CHARS) : all;
}

function textOf(type: AnyType): string {
  if (type instanceof Y.XmlText || type instanceof Y.Text) {
    return (type.toDelta() as { insert: unknown }[]).map((d) => (typeof d.insert === 'string' ? d.insert : '')).join('');
  }
  if (type instanceof Y.XmlElement || type instanceof Y.XmlFragment) {
    const parts = type.toArray().map((child) => textOf(child as AnyType));
    return parts.join(type instanceof Y.XmlFragment ? '\n' : '');
  }
  if (type instanceof Y.Array) return type.toArray().map((v) => (v instanceof Y.AbstractType ? textOf(v) : String(v))).join('\n');
  if (type instanceof Y.Map) return [...type.entries()].map(([k, v]) => `${k}: ${v instanceof Y.AbstractType ? textOf(v) : typeof v === 'string' ? v : JSON.stringify(v)}`).join('\n');
  return '';
}

function showStaleRoomNotice(recovery: string): void {
  if (document.getElementById('room-stale-notice')) return;
  const overlay = document.createElement('div');
  overlay.id = 'room-stale-notice';
  overlay.setAttribute('role', 'alertdialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-labelledby', 'room-stale-text');
  overlay.style.cssText =
    'position:fixed;inset:0;z-index:10000;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.35)';
  const card = document.createElement('div');
  card.style.cssText =
    'max-width:min(560px,92vw);padding:20px 22px;border-radius:10px;background:#fff;color:#222;box-shadow:0 10px 40px rgba(0,0,0,.3);font-size:15px;line-height:1.6';
  const text = document.createElement('p');
  text.id = 'room-stale-text';
  text.style.margin = '0 0 12px';
  const lost = recovery.length > 0;
  text.textContent = lost
    ? '연결이 끊긴 사이 문서가 서버에서 다시 열렸습니다. 연결이 끊긴 뒤 입력한 내용은 문서에 반영되지 않았을 수 있습니다. ' +
      '아래에 이 화면에서 고친 내용을 모아 두었으니 복사하거나 내려받은 뒤 다시 불러와 주세요.'
    : '연결이 끊긴 사이 문서가 서버에서 다시 열렸습니다. 같은 내용이 두 번 들어가지 않도록 최신 문서를 다시 불러옵니다.';
  card.append(text);
  const row = document.createElement('div');
  row.style.cssText = 'display:flex;gap:8px;flex-wrap:wrap;align-items:center';
  const button = (label: string, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = label;
    b.style.cssText = 'padding:6px 14px;border-radius:6px;border:1px solid #888;background:#f5f5f5;cursor:pointer';
    b.addEventListener('click', onClick);
    row.append(b);
    return b;
  };
  let focusTarget: HTMLElement;
  if (lost) {
    const box = document.createElement('textarea');
    box.id = 'room-stale-recovery';
    box.readOnly = true;
    box.value = recovery;
    box.setAttribute('aria-label', '내 입력 내용');
    box.style.cssText = 'display:block;width:100%;box-sizing:border-box;height:140px;margin:0 0 12px;font:13px/1.5 monospace';
    card.append(box);
    const copyBtn = button('내 입력 내용 복사', () => {
      box.select();
      const done = (): void => { copyBtn.textContent = '복사했습니다'; };
      navigator.clipboard?.writeText(recovery).then(done, () => { if (document.execCommand('copy')) done(); }) ??
        (document.execCommand('copy') && done());
    });
    button('내려받기', () => {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([recovery], { type: 'text/plain;charset=utf-8' }));
      a.download = '내 입력 내용.txt';
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    });
    focusTarget = copyBtn;
  }
  const reload = button('다시 불러오기', () => location.reload());
  focusTarget ??= reload;
  card.append(row);
  if (!lost) {
    // 잃을 입력이 없을 때만 보이는 초읽기 뒤 저절로 다시 불러온다
    const count = document.createElement('span');
    count.id = 'room-stale-countdown';
    count.setAttribute('aria-live', 'polite');
    row.append(count);
    let left = AUTO_RELOAD_SECONDS;
    const tick = (): void => {
      count.textContent = `${left}초 뒤 저절로 다시 불러옵니다`;
      if (left-- <= 0) location.reload();
      else setTimeout(tick, 1000);
    };
    tick();
  }
  overlay.append(card);
  document.body.append(overlay);
  focusTarget.focus();
}
