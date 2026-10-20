export interface Permissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  /** 그리드 뷰에서 셀 텍스트 선택·복사 허용 여부 */
  readonly canCopy: boolean;
  /** docx 실시간 동시편집 세션(/ws/edit/:id) 참여 허용 여부 */
  readonly canEdit: boolean;
  /** PDF 편집기(editor-pdf.html)용으로 발급한 토큰 — 토큰 요청에 pdfEditor:true를 보낸 호출자(향기랩)에게만, pdf·내려받기 권한이
   * 있을 때만 붙는다. PDF 편집 방은 이 표시가 있는 토큰만 받는다(8000이 마크업용으로 주는 canEdit 토큰으로 PDF 원본을
   * 고치지 못하게). 없으면 생략돼 기존 토큰 모양은 그대로다. */
  readonly pdfEditor?: boolean;
}

export type DocumentStatus = 'processing' | 'ready' | 'failed';

/** pages: PDF 페이지 이미지 뷰 / grid: 스프레드시트 그리드 뷰 / svg: 도면(dxf/dwg) 벡터 뷰 */
export type ViewKind = 'pages' | 'grid' | 'svg';

export interface DocumentMeta {
  readonly id: string;
  readonly name: string;
  readonly ext: string;
  readonly status: DocumentStatus;
  readonly pageCount: number;
  readonly createdAt: string;
  readonly viewKind?: ViewKind;
  /** 변환 단계 진행률(0~100). 서버가 단계별로 기록한다. */
  readonly progress?: number;
  readonly error?: string;
  /** 변환(PDF)이 끝나기 전에도 편집기를 열 수 있다 — 워드(docx) 편집기는 원본만 읽는다(conversionGate.ts) */
  readonly editReady?: boolean;
}

export interface ViewClaims {
  readonly docId: string;
  readonly viewer: string;
  readonly perms: Permissions;
}
