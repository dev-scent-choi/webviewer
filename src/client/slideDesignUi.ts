/**
 * PowerPoint 편집기의 디자인·전환·애니메이션 탭 — 공유 Y.Map(PPTX_SLIDE_DESIGN_MAP)에 사용자가 바꾼
 * 항목을 쓰고, 원본 슬라이드의 설정(PPTX_SLIDE_DESIGN_ORIGINAL_MAP, 서버가 채움)과 합쳐 화면에 보여준다:
 * 배경 색은 캔버스에 바로, 전환은 레일의 ★과 미리 보기로, 애니메이션은 도형 옆 순서 번호와 미리 보기로.
 */
import type * as Y from 'yjs';
import type { SlideEntry } from '../collab/pptxShared.js';
import {
  ANIMATION_LABELS,
  PPTX_SLIDE_DESIGN_MAP,
  PPTX_SLIDE_DESIGN_ORIGINAL_MAP,
  TRANSITION_LABELS,
  type AnimationEffect,
  type AnimationTrigger,
  type SlideAnimation,
  type SlideDesign,
  type SlideTransition,
  type TransitionKind,
  type TransitionSpeed,
} from '../pptx/slideDesign.js';

export interface SlideDesignHost {
  readonly ydoc: Y.Doc;
  readonly canEdit: boolean;
  readonly canvas: HTMLElement;
  readonly rail: HTMLElement;
  entries(): SlideEntry[];
  activeEntry(): SlideEntry | null;
  /** 마지막으로 고른(커서를 둔) 텍스트 상자·도형의 id */
  selectedShapeId(): string | null;
}

const SPEED_MS: Readonly<Record<TransitionSpeed, number>> = { fast: 500, med: 750, slow: 1000 };

export function setupSlideDesign(host: SlideDesignHost): { onCanvasRendered(): void; onSlidesRendered(): void } {
  const design = host.ydoc.getMap<SlideDesign>(PPTX_SLIDE_DESIGN_MAP);
  const originals = host.ydoc.getMap<Required<SlideDesign>>(PPTX_SLIDE_DESIGN_ORIGINAL_MAP);

  const effective = (entry: SlideEntry): Required<SlideDesign> => {
    const original = entry.sourceSlidePart ? originals.get(entry.sourceSlidePart) : undefined;
    const own = design.get(entry.id) ?? {};
    return {
      background: own.background !== undefined ? own.background : (original?.background ?? null),
      transition: own.transition !== undefined ? own.transition : (original?.transition ?? null),
      animations: own.animations ?? original?.animations ?? [],
    };
  };
  const patch = (entry: SlideEntry, change: Partial<SlideDesign>) => {
    const next: Record<string, unknown> = { ...(design.get(entry.id) ?? {}), ...change };
    for (const key of Object.keys(change)) if ((change as Record<string, unknown>)[key] === undefined) delete next[key];
    design.set(entry.id, next as SlideDesign);
  };

  // ---------- 디자인: 배경 서식 ----------
  let lastBackground = '#ffffff';
  for (const swatch of document.querySelectorAll<HTMLElement>('[data-slide-bg]')) {
    swatch.addEventListener('click', () => {
      const entry = host.activeEntry();
      if (!entry || !host.canEdit) return;
      lastBackground = swatch.dataset.slideBg ?? '#ffffff';
      patch(entry, { background: lastBackground });
    });
  }
  const customBg = document.getElementById('slide-bg-custom') as HTMLInputElement | null;
  customBg?.addEventListener('input', () => {
    const entry = host.activeEntry();
    if (!entry || !host.canEdit) return;
    lastBackground = customBg.value;
    patch(entry, { background: lastBackground });
  });
  document.getElementById('btn-bg-apply-all')?.addEventListener('click', () => {
    const entry = host.activeEntry();
    if (!entry || !host.canEdit) return;
    const color = effective(entry).background;
    host.ydoc.transact(() => {
      for (const e of host.entries()) patch(e, { background: color ?? undefined });
    });
  });
  document.getElementById('btn-bg-reset')?.addEventListener('click', () => {
    const entry = host.activeEntry();
    if (entry && host.canEdit) patch(entry, { background: undefined });
  });

  // ---------- 전환 ----------
  const speedSelect = document.getElementById('transition-speed') as HTMLSelectElement | null;
  for (const button of document.querySelectorAll<HTMLElement>('[data-transition]')) {
    button.addEventListener('click', () => {
      const entry = host.activeEntry();
      if (!entry || !host.canEdit) return;
      const kind = button.dataset.transition as TransitionKind | 'none';
      patch(entry, { transition: kind === 'none' ? null : { kind, speed: (speedSelect?.value as TransitionSpeed) ?? 'med' } });
      if (kind !== 'none') previewTransition();
    });
  }
  speedSelect?.addEventListener('change', () => {
    const entry = host.activeEntry();
    const current = entry ? effective(entry).transition : null;
    if (entry && current && host.canEdit) patch(entry, { transition: { ...current, speed: speedSelect.value as TransitionSpeed } });
  });
  document.getElementById('btn-transition-apply-all')?.addEventListener('click', () => {
    const entry = host.activeEntry();
    if (!entry || !host.canEdit) return;
    const transition = effective(entry).transition;
    host.ydoc.transact(() => {
      for (const e of host.entries()) patch(e, { transition });
    });
  });
  document.getElementById('btn-transition-preview')?.addEventListener('click', previewTransition);

  function previewTransition(): void {
    const entry = host.activeEntry();
    const transition: SlideTransition | null = entry ? effective(entry).transition : null;
    if (!transition) return;
    const keyframes: Readonly<Record<TransitionKind, Keyframe[]>> = {
      fade: [{ opacity: 0 }, { opacity: 1 }],
      push: [{ transform: 'translateY(100%)' }, { transform: 'translateY(0)' }],
      wipe: [{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0 0 0 0)' }],
      split: [{ clipPath: 'inset(50% 0 50% 0)' }, { clipPath: 'inset(0 0 0 0)' }],
      cover: [{ transform: 'translateX(100%)' }, { transform: 'translateX(0)' }],
      zoom: [{ transform: 'scale(0.3)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }],
    };
    host.canvas.animate(keyframes[transition.kind], { duration: SPEED_MS[transition.speed], easing: 'ease-out' });
  }

  // ---------- 애니메이션 ----------
  const triggerSelect = document.getElementById('animation-trigger') as HTMLSelectElement | null;
  for (const button of document.querySelectorAll<HTMLElement>('[data-animation]')) {
    button.addEventListener('click', () => {
      const entry = host.activeEntry();
      const shapeId = host.selectedShapeId();
      if (!entry || !host.canEdit) return;
      if (!shapeId) {
        window.alert('애니메이션을 줄 텍스트 상자나 도형을 먼저 클릭하세요.');
        return;
      }
      const effect = button.dataset.animation as AnimationEffect | 'none';
      const list = effective(entry).animations.filter((a) => a.shapeId !== shapeId);
      const existing = effective(entry).animations.find((a) => a.shapeId === shapeId);
      const next: SlideAnimation[] =
        effect === 'none'
          ? list
          : existing
            ? effective(entry).animations.map((a) => (a.shapeId === shapeId ? { ...a, effect } : a))
            : [...list, { shapeId, effect, trigger: (triggerSelect?.value as AnimationTrigger) ?? 'click' }];
      patch(entry, { animations: next });
      if (effect !== 'none') void previewAnimations(shapeId);
    });
  }
  triggerSelect?.addEventListener('change', () => {
    const entry = host.activeEntry();
    const shapeId = host.selectedShapeId();
    if (!entry || !shapeId || !host.canEdit) return;
    const list = effective(entry).animations;
    if (!list.some((a) => a.shapeId === shapeId)) return;
    patch(entry, { animations: list.map((a) => (a.shapeId === shapeId ? { ...a, trigger: triggerSelect.value as AnimationTrigger } : a)) });
  });
  document.getElementById('btn-animation-preview')?.addEventListener('click', () => void previewAnimations(null));

  const animationKeyframes: Readonly<Record<AnimationEffect, Keyframe[]>> = {
    appear: [{ opacity: 1 }, { opacity: 1 }],
    fade: [{ opacity: 0 }, { opacity: 1 }],
    fly: [{ transform: 'translateY(120%)', opacity: 0.4 }, { transform: 'translateY(0)', opacity: 1 }],
    wipe: [{ clipPath: 'inset(100% 0 0 0)' }, { clipPath: 'inset(0 0 0 0)' }],
  };
  /** 미리 보기 — 숨겼다가 순서대로 나타낸다(클릭 묶음 사이는 잠깐 쉰다). onlyShape면 그 도형만. */
  async function previewAnimations(onlyShape: string | null): Promise<void> {
    const entry = host.activeEntry();
    if (!entry) return;
    const list = effective(entry).animations.filter((a) => !onlyShape || a.shapeId === onlyShape);
    const boxOf = (id: string) => host.canvas.querySelector<HTMLElement>(`.slide-textbox[data-shape-id="${CSS.escape(id)}"]`);
    for (const a of list) {
      const box = boxOf(a.shapeId);
      if (box) box.style.visibility = 'hidden';
    }
    const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
    for (const [i, a] of list.entries()) {
      if (i > 0 && a.trigger === 'click') await wait(450);
      if (a.trigger === 'after') await wait(500);
      const box = boxOf(a.shapeId);
      if (!box) continue;
      box.style.visibility = '';
      box.animate(animationKeyframes[a.effect], { duration: a.effect === 'appear' ? 1 : 500, easing: 'ease-out' });
    }
  }

  function syncRibbon(): void {
    const entry = host.activeEntry();
    const eff = entry ? effective(entry) : null;
    for (const button of document.querySelectorAll<HTMLElement>('[data-transition]')) {
      const kind = button.dataset.transition;
      button.setAttribute('aria-pressed', String(eff ? (eff.transition ? eff.transition.kind === kind : kind === 'none') : false));
    }
    if (speedSelect && eff?.transition) speedSelect.value = eff.transition.speed;
    const shapeId = host.selectedShapeId();
    const anim = eff && shapeId ? eff.animations.find((a) => a.shapeId === shapeId) : undefined;
    for (const button of document.querySelectorAll<HTMLElement>('[data-animation]')) {
      const effect = button.dataset.animation;
      button.setAttribute('aria-pressed', String(shapeId ? (anim ? anim.effect === effect : effect === 'none') : false));
    }
    if (triggerSelect && anim) triggerSelect.value = anim.trigger;
    const label = document.getElementById('animation-target');
    if (label) {
      label.textContent = shapeId
        ? anim
          ? `선택한 개체: ${ANIMATION_LABELS[anim.effect]}`
          : '선택한 개체: 애니메이션 없음'
        : '텍스트 상자나 도형을 클릭하세요';
    }
  }

  function onCanvasRendered(): void {
    const entry = host.activeEntry();
    if (!entry) return;
    const eff = effective(entry);
    const own = design.get(entry.id);
    host.canvas.querySelector('.slide-bg-override')?.remove();
    for (const el of host.canvas.querySelectorAll<HTMLElement>('[data-background]')) el.hidden = own?.background !== undefined;
    if (own?.background) {
      const bg = document.createElement('div');
      bg.className = 'slide-bg-override';
      bg.style.background = own.background;
      host.canvas.prepend(bg);
    }
    for (const badge of host.canvas.querySelectorAll('.slide-anim-badge')) badge.remove();
    let number = 0;
    for (const a of eff.animations) {
      if (a.trigger === 'click' || number === 0) number += 1;
      const box = host.canvas.querySelector<HTMLElement>(`.slide-textbox[data-shape-id="${CSS.escape(a.shapeId)}"]`);
      if (!box) continue;
      const badge = document.createElement('span');
      badge.className = 'slide-anim-badge';
      badge.textContent = String(number);
      badge.title = `${number}번째 · ${ANIMATION_LABELS[a.effect]} · ${a.trigger === 'click' ? '클릭할 때' : a.trigger === 'with' ? '이전 효과와 함께' : '이전 효과 다음에'}`;
      box.append(badge);
    }
    syncRibbon();
  }

  function onSlidesRendered(): void {
    const entries = host.entries();
    for (const card of host.rail.querySelectorAll<HTMLElement>('.slide-card')) {
      card.querySelector('.slide-card-transition')?.remove();
      const entry = entries.find((e) => e.id === card.dataset.id);
      if (!entry) continue;
      const eff = effective(entry);
      if (!eff.transition && eff.animations.length === 0) continue;
      const star = document.createElement('span');
      star.className = 'slide-card-transition';
      star.textContent = '★';
      star.title = [eff.transition ? `전환: ${TRANSITION_LABELS[eff.transition.kind]}` : '', eff.animations.length ? `애니메이션 ${eff.animations.length}개` : '']
        .filter(Boolean)
        .join(' · ');
      card.querySelector('.slide-card-thumb')?.append(star);
    }
  }

  design.observe(() => {
    onCanvasRendered();
    onSlidesRendered();
  });
  originals.observe(() => {
    onCanvasRendered();
    onSlidesRendered();
  });
  document.addEventListener('focusin', () => syncRibbon());
  return { onCanvasRendered, onSlidesRendered };
}
