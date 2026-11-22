// scripts/ribbonGenLib.mjs 타입 — 함수가 바뀌면 같이 고친다
export interface RibbonControl {
  key: string; name: string; size: 'big' | 'small';
  kind: 'button' | 'split' | 'dropdown' | 'gallery' | 'toggle' | 'input' | 'select' | 'info';
  pending?: true; ref?: string; unit?: string; proxy?: string; map?: 'same' | 'moved' | 'nearest' | 'proxy';
  officeKind?: string; convert?: 'big→small' | 'small→big'; fix?: string; note?: string;
  row?: number; iconOnly?: true; break?: true; stackTitle?: string; icon?: string; items?: string[]; unverified?: true;
}
export interface RibbonGroup { id: string; name: string; collapse: number; layout: 'columns' | 'rows' | 'fields'; launcher: null | { pending: string } | { proxy: string }; domId?: string; className?: string; note?: string; unverified?: true; controls: RibbonControl[] }
export interface RibbonTab { id: string; name: string; editOnly: boolean; panelId?: string; note?: string; trigger?: string; wired?: string | null; unverified?: true; groups: RibbonGroup[] }
export interface RibbonApp { page: string; backstage: { office: string[]; existing: string[] }; tabs: RibbonTab[]; contextual: RibbonTab[] }
export type AppName = 'word' | 'excel' | 'powerpoint';
export interface RibbonInventory { version: number; spec: string; source: string; schema: Record<string, string>; summary: Record<AppName, Record<string, number>>; apps: Record<AppName, RibbonApp> }
export declare const INVENTORY_PATH: string;
export declare const APP_OF_EXT: Readonly<Record<'docx' | 'xlsx' | 'pptx', AppName>>;
export declare const KINDS: readonly string[];
export declare function loadInventory(file?: string): RibbonInventory;
export declare function controlKey(tabId: string, groupId: string, c: RibbonControl): string;
export declare function tabCount(app: RibbonApp): number;
export declare function eachControl(app: RibbonApp): Generator<{ tab: RibbonTab; group: RibbonGroup; c: RibbonControl; key: string }>;
export declare function pendingIn(app: RibbonApp, tabId: string): number;
export declare function launcherCount(app: RibbonApp): number;
export declare function summarize(app: RibbonApp): Record<string, number>;
export declare function validateInventory(inv: RibbonInventory): string[];
export declare function serializeInventory(inv: RibbonInventory): string;
export declare function requiredIcons(inv: RibbonInventory): string[];
