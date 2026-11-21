/** 통합 보기 껍데기의 공유 타입 — 설계 §4 ViewBody(계획의 "설계와 맞춘 이름" 참고) */
export interface PageSize {
  readonly w: number;
  readonly h: number;
}

/** 쪽 크기 대비 비율(0~1) 사각형 — 서버 찾기(/api/view/search)의 rects와 같은 모양 */
export interface NormRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export interface SearchHit {
  readonly page: number;
  readonly rects: readonly NormRect[];
}

export interface ViewPermissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
  readonly canEdit: boolean;
  readonly pdfEditor?: boolean;
}

/** /api/view/info의 artcraft 블록(서버 src/artcraftEngine.ts ArtcraftInfo와 같은 모양) */
export interface ArtcraftInfo {
  readonly kind: 'word' | 'deck';
  readonly ext: string;
  readonly experimental: boolean;
  readonly originalUrl: string;
  readonly assets: { readonly wasm: string; readonly fonts: readonly string[] };
}

/** GET /api/view/info 의 data(이 화면이 쓰는 것만) */
export interface ViewInfo {
  readonly docId: string;
  readonly name: string;
  readonly ext: string;
  readonly pageCount: number;
  readonly viewKind: 'pages' | 'grid' | 'svg';
  readonly pages?: readonly PageSize[];
  readonly permissions: ViewPermissions;
  readonly viewer: string;
  /** 편집기 주소(토큰 포함) — 없으면 편집 단추가 없다(설계 V6) */
  readonly editorUrl?: string;
  /** 이 문서의 보기 화면 주소(토큰 포함) — 쪽 문서는 /view.html, 엑셀·도면·한글은 /viewer.html */
  readonly viewerUrl: string;
  /** 첫 변환 중(editReady) — 쪽 자리에 "쪽을 준비하는 중"(설계 V8) */
  readonly processing: boolean;
  /** 여는 엔진 — artcraft면 브라우저 wasm 엔진(artcraft 블록이 함께 온다) */
  readonly engine?: 'libreoffice' | 'artcraft';
  /** 아트크래프트를 요청했지만 서버 변환으로 여는 이유(한 줄 안내로 보인다) */
  readonly engineReason?: string;
  readonly artcraft?: ArtcraftInfo;
}

export type BodyKind = 'pages-png' | 'pages-pdfjs' | 'pages-artcraft' | 'grid' | 'cad' | 'frame';

/** number = 확대 %(logic.ts ZOOM_STEPS, 100 = 쪽 칸 폭 전체) */
export type ZoomMode = number | 'fit-width' | 'fit-page';

export interface ViewBody {
  readonly kind: BodyKind;
  readonly textSelectable: boolean;
  mount(host: HTMLElement): Promise<void>;
  /** 쪽 개념이 없으면 0(막대의 쪽 칸 숨김) */
  pageCount(): number;
  goto(page: number): void;
  onPage(cb: (page: number) => void): void;
  zoom(z: ZoomMode): void;
  rotate(deg: 90 | -90): void;
  search?(q: string, signal: AbortSignal): Promise<SearchHit[]>;
  thumbnail?(page: number, width: number): Promise<HTMLCanvasElement | HTMLImageElement>;
  destroy(): void;
}
