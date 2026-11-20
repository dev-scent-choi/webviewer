import type * as Y from 'yjs';
import type { WebsocketProvider } from 'y-websocket';
import * as decoding from 'lib0/decoding';
import type * as encoding from 'lib0/encoding';
import { MESSAGE_ROOM_EPOCH, MESSAGE_ROOM_STALE, ROOM_EPOCH_PARAM } from '../collab/roomEpoch.js';

type Handler = (encoder: encoding.Encoder, decoder: decoding.Decoder) => void;

/**
 * 편집기(워드·엑셀·PPT·PDF) 공용 — 방 세대 확인(src/collab/roomEpoch.ts 설명). 서버가 알려 준 세대를 다음 접속 주소에 싣고,
 * 이전 세대로 다시 접속했다고 거절당하면 다시 접속하지 않고(옛 상태를 새 방에 합치지 않는다) 안내 후 새로 불러온다.
 * 끊긴 동안 이 화면에서 고친 내용이 있었으면 그 사실을 안내에 적는다 — 그 편집은 새 방의 문서와 안전하게 맞출 방법이 없어
 * 다시 보내지 않는다(옛 세대의 Yjs 변경을 새 세대 문서에 다시 적용하면 바로 이 중복이 생긴다).
 */
export function guardRoomEpoch(provider: WebsocketProvider, ydoc: Y.Doc): void {
  let editedWhileOffline = false;
  ydoc.on('update', (_update: Uint8Array, origin: unknown) => {
    if (origin !== provider && !provider.wsconnected) editedWhileOffline = true;
  });
  const handlers = provider.messageHandlers as unknown as Handler[];
  handlers[MESSAGE_ROOM_EPOCH] = (_encoder, decoder) => {
    provider.params[ROOM_EPOCH_PARAM] = decoding.readVarString(decoder);
    editedWhileOffline = false; // 이 세대에 붙은 뒤의 편집은 정상적으로 동기화된다
  };
  handlers[MESSAGE_ROOM_STALE] = () => {
    provider.shouldConnect = false;
    provider.disconnect();
    showStaleRoomNotice(editedWhileOffline);
  };
}

function showStaleRoomNotice(editedWhileOffline: boolean): void {
  if (document.getElementById('room-stale-notice')) return;
  const overlay = document.createElement('div');
  overlay.id = 'room-stale-notice';
  overlay.setAttribute('role', 'alertdialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.style.cssText =
    'position:fixed;inset:0;z-index:10000;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.35)';
  const card = document.createElement('div');
  card.style.cssText =
    'max-width:min(480px,90vw);padding:20px 22px;border-radius:10px;background:#fff;color:#222;box-shadow:0 10px 40px rgba(0,0,0,.3);font-size:15px;line-height:1.6';
  const text = document.createElement('p');
  text.style.margin = '0 0 14px';
  text.textContent =
    '연결이 끊긴 사이 문서가 서버에서 다시 열렸습니다. 같은 내용이 두 번 들어가지 않도록 최신 문서를 다시 불러옵니다.' +
    (editedWhileOffline ? ' 연결이 끊긴 동안 입력한 내용은 반영되지 않았습니다 — 다시 불러온 뒤 다시 입력해 주세요.' : '');
  const button = document.createElement('button');
  button.type = 'button';
  button.textContent = '다시 불러오기';
  button.style.cssText = 'padding:6px 16px;border-radius:6px;border:1px solid #888;background:#f5f5f5;cursor:pointer';
  button.addEventListener('click', () => location.reload());
  card.append(text, button);
  overlay.append(card);
  document.body.append(overlay);
  // 끊긴 동안 입력한 내용이 없으면 잠깐 안내만 하고 저절로 다시 불러온다(입력이 있었으면 읽을 시간을 주고 사용자가 누른다)
  if (!editedWhileOffline) setTimeout(() => location.reload(), 2500);
  button.focus();
}
