import { describe, expect, it } from 'vitest';
import { paginateFlow, type FlowItem } from '../src/client/docxLayout.js';

const para = (height: number, extra: Partial<FlowItem> = {}): FlowItem => ({
  height, marginTop: 0, marginBottom: 0, kind: 'para', lineHeight: 20, notes: 0, section: 0, breakBefore: false, breakAfter: false, ...extra,
});
const body = () => 100;

describe('넘침 쪽 나눔(높이표, m6)', () => {
  it('넘치는 블록 앞에서 쪽이 바뀐다 — 여백은 겹쳐 큰 쪽만', () => {
    expect(paginateFlow([para(40, { marginBottom: 10 }), para(40, { marginTop: 6, kind: 'block' }), para(20, { kind: 'block' })], body)).toEqual([{ index: 2, offset: 0 }]);
  });
  it('각주 높이만큼 쪽 본문이 줄어든다', () => {
    expect(paginateFlow([para(40, { kind: 'block' }), para(40, { kind: 'block', notes: 30 })], body)).toEqual([{ index: 1, offset: 0 }]);
    expect(paginateFlow([para(40, { kind: 'block', notes: 30 }), para(40, { kind: 'block' })], body)).toEqual([{ index: 1, offset: 0 }]);
  });
  it('긴 문단은 줄 단위로 나누고 앞뒤 2줄은 남긴다', () => {
    // 60px 쓴 쪽에 10줄(200px) 문단: 남은 40px = 2줄 → 2줄 두고, 나머지 8줄 160px는 다음 쪽 100px(5줄)에 → 또 나눔
    expect(paginateFlow([para(60, { kind: 'block' }), para(200)], body)).toEqual([{ index: 1, offset: 40 }, { index: 1, offset: 140 }]);
    // 3줄 문단은 2+1로 나눌 수 없다 — 통째로 다음 쪽
    expect(paginateFlow([para(60, { kind: 'block' }), para(60)], body)).toEqual([{ index: 1, offset: 0 }]);
  });
  it('쪽보다 긴 표는 높이로 이어 센다, 그림(block)은 나누지 않는다', () => {
    expect(paginateFlow([para(250, { kind: 'split' })], body)).toEqual([{ index: 0, offset: 100 }, { index: 0, offset: 200 }]);
    expect(paginateFlow([para(250, { kind: 'block' })], body)).toEqual([]);
  });
  it('명시적 쪽 나눔·구역 뒤에서 다시 센다, 아직 안 잰 블록은 건너뛴다', () => {
    expect(paginateFlow([para(90, { kind: 'block', breakAfter: true }), para(90, { kind: 'block' }), para(null as unknown as number), para(90, { kind: 'block', breakBefore: true })], body)).toEqual([]);
  });
});
