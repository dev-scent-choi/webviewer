export interface Permissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
}

export type DocumentStatus = 'processing' | 'ready' | 'failed';

export interface DocumentMeta {
  readonly id: string;
  readonly name: string;
  readonly ext: string;
  readonly status: DocumentStatus;
  readonly pageCount: number;
  readonly createdAt: string;
  /** 변환 단계 진행률(0~100). 서버가 단계별로 기록한다. */
  readonly progress?: number;
  readonly error?: string;
}

export interface ViewClaims {
  readonly docId: string;
  readonly viewer: string;
  readonly perms: Permissions;
}
