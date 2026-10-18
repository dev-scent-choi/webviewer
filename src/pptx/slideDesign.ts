/**
 * PowerPoint 디자인·전환·애니메이션 탭의 슬라이드 단위 설정 — 배경 색, 화면 전환, 개체 나타내기
 * 애니메이션. 편집기는 공유 Y.Map(PPTX_SLIDE_DESIGN_MAP, 슬라이드 항목 id → SlideDesign)에 사용자가
 * 바꾼 항목만 담고, 저장할 때 그 슬라이드 XML의 <p:bg>·<p:transition>·<p:timing>만 바꾼다. 바꾸지
 * 않은 항목(undefined)은 원본 XML을 그대로 둔다. 이 파일은 브라우저 번들에도 들어간다.
 */

export type TransitionKind = 'fade' | 'push' | 'wipe' | 'split' | 'cover' | 'zoom';
export type TransitionSpeed = 'fast' | 'med' | 'slow';
export type AnimationEffect = 'appear' | 'fade' | 'fly' | 'wipe';
export type AnimationTrigger = 'click' | 'with' | 'after';

export interface SlideTransition {
  readonly kind: TransitionKind;
  readonly speed: TransitionSpeed;
}

export interface SlideAnimation {
  /** 도형의 cNvPr id(슬라이드 XML 안) */
  readonly shapeId: string;
  readonly effect: AnimationEffect;
  readonly trigger: AnimationTrigger;
}

export interface SlideDesign {
  /** '#rrggbb' 단색 배경, null이면 슬라이드 자체 배경을 지워 레이아웃·마스터 배경을 쓴다 */
  readonly background?: string | null;
  /** null이면 전환 없음 */
  readonly transition?: SlideTransition | null;
  /** 있으면 이 목록으로 애니메이션 전체를 다시 쓴다(빈 배열 = 애니메이션 없음) */
  readonly animations?: readonly SlideAnimation[];
}

export const PPTX_SLIDE_DESIGN_MAP = 'slideDesign';
/** 원본 슬라이드에서 읽은 현재 설정(편집기 표시용, 서버가 시딩) */
export const PPTX_SLIDE_DESIGN_ORIGINAL_MAP = 'slideDesignOriginal';

export const TRANSITION_LABELS: Readonly<Record<TransitionKind, string>> = {
  fade: '밝기 변화',
  push: '밀어내기',
  wipe: '닦아내기',
  split: '나누기',
  cover: '덮기',
  zoom: '확대/축소',
};
export const ANIMATION_LABELS: Readonly<Record<AnimationEffect, string>> = {
  appear: '나타내기',
  fade: '밝기 변화',
  fly: '날아오기',
  wipe: '닦아내기',
};
const PRESET_ID: Readonly<Record<AnimationEffect, number>> = { appear: 1, fly: 2, fade: 10, wipe: 22 };
const EFFECT_BY_PRESET: Readonly<Record<number, AnimationEffect>> = { 1: 'appear', 2: 'fly', 10: 'fade', 22: 'wipe' };
export const ANIMATION_DURATION_MS: Readonly<Record<AnimationEffect, number>> = { appear: 1, fade: 500, fly: 500, wipe: 500 };

// ---------------------------------------------------------------------------
// 원본 슬라이드 XML 읽기
// ---------------------------------------------------------------------------

export function readSlideDesign(slideXml: string): Required<SlideDesign> {
  const bgColor = /<p:bg>[\s\S]*?<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(slideXml)?.[1];
  // PowerPoint 2010+는 전환을 mc:AlternateContent로 감싸 p14 요소와 기본 요소를 함께 쓴다 — 어느 쪽이든 종류만 본다
  const transitionXml = /<p:transition\b[^>]*?(?:\/>|>([\s\S]*?)<\/p:transition>)/.exec(slideXml);
  let transition: SlideTransition | null = null;
  if (transitionXml) {
    const kind = (['fade', 'push', 'wipe', 'split', 'cover', 'zoom'] as const).find((k) => new RegExp(`<p(?:14)?:${k}\\b`).test(transitionXml[1] ?? '')) ?? 'fade';
    const spd = /\bspd="(slow|med|fast)"/.exec(transitionXml[0])?.[1] as TransitionSpeed | undefined;
    transition = transitionXml[1] !== undefined && transitionXml[1].trim() !== '' ? { kind, speed: spd ?? 'med' } : null;
  }
  const animations: SlideAnimation[] = [];
  const timing = /<p:timing>([\s\S]*?)<\/p:timing>/.exec(slideXml)?.[1] ?? '';
  // 효과 노드 안의 대상(spTgt)은 중첩된 <p:cTn> 뒤에 나오므로, 이 효과 시작부터 다음 효과 시작까지에서 찾는다
  const effects = [...timing.matchAll(/<p:cTn\b([^>]*presetClass="(\w+)"[^>]*)>/g)];
  for (const [i, m] of effects.entries()) {
    if (m[2] !== 'entr') continue;
    const attrs = m[1];
    const presetId = Number(/presetID="(\d+)"/.exec(attrs)?.[1] ?? 0);
    const nodeType = /nodeType="(\w+)"/.exec(attrs)?.[1] ?? 'clickEffect';
    const end = effects[i + 1]?.index ?? timing.length;
    const spid = /<p:spTgt spid="(\d+)"/.exec(timing.slice((m.index ?? 0) + m[0].length, end))?.[1];
    if (!spid) continue;
    animations.push({
      shapeId: spid,
      effect: EFFECT_BY_PRESET[presetId] ?? 'fade',
      trigger: nodeType === 'withEffect' ? 'with' : nodeType === 'afterEffect' ? 'after' : 'click',
    });
  }
  return { background: bgColor ? `#${bgColor.toLowerCase()}` : null, transition, animations };
}

// ---------------------------------------------------------------------------
// 슬라이드 XML에 반영(저장)
// ---------------------------------------------------------------------------

function applyBackground(xml: string, background: string | null): string {
  const withoutBg = xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/, '');
  if (background === null) return withoutBg;
  const hex = background.replace('#', '').toUpperCase();
  const bg = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
  return withoutBg.replace(/(<p:cSld\b[^>]*>)/, `$1${bg}`);
}

const TRANSITION_XML: Readonly<Record<TransitionKind, string>> = {
  fade: '<p:fade/>',
  push: '<p:push dir="u"/>',
  wipe: '<p:wipe dir="r"/>',
  split: '<p:split orient="horz" dir="out"/>',
  cover: '<p:cover dir="l"/>',
  zoom: '<p:zoom/>',
};

/** 기존 전환(mc:AlternateContent로 감싼 것 포함)과 애니메이션(<p:timing>)을 지운다 */
function removeTransition(xml: string): string {
  return xml
    .replace(/<mc:AlternateContent\b[^>]*>(?:(?!<\/mc:AlternateContent>)[\s\S])*?<p:transition\b[\s\S]*?<\/mc:AlternateContent>/g, '')
    .replace(/<p:transition\b[^>]*\/>|<p:transition\b[^>]*>[\s\S]*?<\/p:transition>/g, '');
}

const TRANSITION_ELEMENT =
  /^\s*(?:<p:transition\b[^>]*\/>|<p:transition\b[^>]*>[\s\S]*?<\/p:transition>|<mc:AlternateContent\b[^>]*>(?:(?!<\/mc:AlternateContent>)[\s\S])*?<p:transition\b[\s\S]*?<\/mc:AlternateContent>)/;

/** CT_Slide 순서(cSld, clrMapOvr, transition, timing, extLst)대로 전환·애니메이션 자리에 넣는다 */
function insertSlideChild(xml: string, fragment: string, kind: 'transition' | 'timing'): string {
  const head = /<p:clrMapOvr\b[^>]*\/>|<p:clrMapOvr\b[^>]*>[\s\S]*?<\/p:clrMapOvr>/.exec(xml) ?? /<\/p:cSld>/.exec(xml);
  if (!head) return xml;
  let at = head.index + head[0].length;
  if (kind === 'timing') {
    const transition = TRANSITION_ELEMENT.exec(xml.slice(at));
    if (transition) at += transition[0].length;
  }
  return xml.slice(0, at) + fragment + xml.slice(at);
}

function applyTransition(xml: string, transition: SlideTransition | null): string {
  const cleared = removeTransition(xml);
  if (!transition) return cleared;
  return insertSlideChild(cleared, `<p:transition spd="${transition.speed}">${TRANSITION_XML[transition.kind]}</p:transition>`, 'transition');
}

/** 개체 하나의 나타내기 효과 노드(<p:par> … presetClass="entr") */
function effectXml(anim: SlideAnimation, nextId: () => number, delayMs: number): string {
  const nodeType = anim.trigger === 'with' ? 'withEffect' : anim.trigger === 'after' ? 'afterEffect' : 'clickEffect';
  const target = `<p:tgtEl><p:spTgt spid="${anim.shapeId}"/></p:tgtEl>`;
  const presetSubtype = anim.effect === 'fly' || anim.effect === 'wipe' ? 4 : 0;
  const setVisible =
    `<p:set><p:cBhvr><p:cTn id="${nextId()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${target}` +
    '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>';
  let body = setVisible;
  if (anim.effect === 'fade') {
    body += `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${nextId()}" dur="500"/>${target}</p:cBhvr></p:animEffect>`;
  } else if (anim.effect === 'wipe') {
    body += `<p:animEffect transition="in" filter="wipe(up)"><p:cBhvr><p:cTn id="${nextId()}" dur="500"/>${target}</p:cBhvr></p:animEffect>`;
  } else if (anim.effect === 'fly') {
    const anim1 = (attr: string, from: string, to: string) =>
      `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${nextId()}" dur="500" fill="hold"/>${target}` +
      `<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>` +
      `<p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
    body += anim1('ppt_x', '#ppt_x', '#ppt_x') + anim1('ppt_y', '1+#ppt_h/2', '#ppt_y');
  }
  return (
    `<p:par><p:cTn id="${nextId()}" presetID="${PRESET_ID[anim.effect]}" presetClass="entr" presetSubtype="${presetSubtype}" fill="hold" grpId="0" nodeType="${nodeType}">` +
    `<p:stCondLst><p:cond delay="${delayMs}"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>`
  );
}

/**
 * PowerPoint 표준 구조의 <p:timing>: 주 시퀀스 → 클릭 그룹(par, delay=indefinite) → 같은 순간 묶음(par,
 * delay=0 또는 앞 효과가 끝난 시점) → 효과. "이전 효과와 함께"는 같은 묶음에, "이전 효과 다음에"는 같은
 * 클릭 그룹의 다음 묶음에 들어간다. 첫 효과가 함께/다음이면 슬라이드가 열릴 때 바로 시작한다.
 */
export function buildTimingXml(animations: readonly SlideAnimation[]): string {
  if (animations.length === 0) return '';
  let id = 2;
  const nextId = () => {
    id += 1;
    return id;
  };
  interface Moment {
    effects: SlideAnimation[];
    delay: number;
  }
  const clickGroups: Moment[][] = [];
  for (const anim of animations) {
    const group = clickGroups[clickGroups.length - 1];
    if (anim.trigger === 'click' || !group) {
      clickGroups.push([{ effects: [anim], delay: 0 }]);
    } else if (anim.trigger === 'with') {
      group[group.length - 1].effects.push(anim);
    } else {
      const last = group[group.length - 1];
      const lastEnd = last.delay + Math.max(...last.effects.map((e) => ANIMATION_DURATION_MS[e.effect]));
      group.push({ effects: [anim], delay: lastEnd });
    }
  }
  const groupsXml = clickGroups
    .map((moments, gi) => {
      const firstTrigger = moments[0].effects[0].trigger;
      const groupDelay = gi === 0 && firstTrigger !== 'click' ? '0' : 'indefinite';
      const momentsXml = moments
        .map((moment) => {
          const cTnId = nextId();
          const effects = moment.effects.map((e) => effectXml(e, nextId, 0)).join('');
          return `<p:par><p:cTn id="${cTnId}" fill="hold"><p:stCondLst><p:cond delay="${moment.delay}"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par>`;
        })
        .join('');
      const groupId = nextId();
      return `<p:par><p:cTn id="${groupId}" fill="hold"><p:stCondLst><p:cond delay="${groupDelay}"/></p:stCondLst><p:childTnLst>${momentsXml}</p:childTnLst></p:cTn></p:par>`;
    })
    .join('');
  const shapes = [...new Set(animations.map((a) => a.shapeId))];
  return (
    '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groupsXml}</p:childTnLst></p:cTn>` +
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>' +
    '</p:childTnLst></p:cTn></p:par></p:tnLst>' +
    `<p:bldLst>${shapes.map((s) => `<p:bldP spid="${s}" grpId="0" animBg="1"/>`).join('')}</p:bldLst></p:timing>`
  );
}

function applyAnimations(xml: string, animations: readonly SlideAnimation[]): string {
  const cleared = xml.replace(/<p:timing>[\s\S]*?<\/p:timing>/, '');
  const timing = buildTimingXml(animations);
  return timing ? insertSlideChild(cleared, timing, 'timing') : cleared;
}

/** 사용자가 바꾼 항목만 슬라이드 XML에 반영한다 */
export function applySlideDesign(slideXml: string, design: SlideDesign): string {
  let xml = slideXml;
  if (design.background !== undefined) xml = applyBackground(xml, design.background);
  if (design.transition !== undefined) xml = applyTransition(xml, design.transition);
  if (design.animations !== undefined) xml = applyAnimations(xml, design.animations);
  return xml;
}
