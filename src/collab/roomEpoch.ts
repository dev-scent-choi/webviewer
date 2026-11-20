/**
 * 방 세대(epoch) — 동시편집 방(room.ts)이 새로 만들어질 때마다 새 값. 웹소켓이 끊겼다(와이파이·노트북 덮개·서버 재시작)
 * 다시 이어질 때, 클라이언트가 이전 세대 방에서 받은 Yjs 상태를 새로 시딩된 방에 합치면 슬라이드·문단·행이 두 벌이 되고
 * (pptx: 슬라이드 3 → 6, presentation.xml에 같은 sldId) 그대로 저장됐다(2차 검토 RX-C1, 운영 94ab62b에서도 재현).
 *
 * 약속: 서버는 연결을 받으면 MESSAGE_ROOM_EPOCH로 지금 방의 세대를 알려 준다. 클라이언트는 그 값을 웹소켓 주소의
 * `epoch` 매개변수로 들고 다시 접속한다. 서버는 그 값이 지금 방과 다르면(또는 방이 없으면) 동기화 메시지를 하나도 받지
 * 않고 MESSAGE_ROOM_STALE를 보낸 뒤 연결을 닫는다. 클라이언트는 다시 접속하지 않고 안내 후 새로 불러온다.
 * 서버(room.ts)와 편집기(클라이언트 번들) 양쪽이 이 파일을 쓴다 — DOM·서버 전용 의존성 없음.
 */
/** 0~6은 y-websocket 표준과 기존 사용자 정의 메시지(room.ts)가 쓴다 */
export const MESSAGE_ROOM_EPOCH = 7;
export const MESSAGE_ROOM_STALE = 8;
/** 웹소켓 주소 매개변수 이름 */
export const ROOM_EPOCH_PARAM = 'epoch';
/** 이전 세대로 다시 접속한 연결을 닫는 코드 */
export const ROOM_STALE_CLOSE_CODE = 4409;
