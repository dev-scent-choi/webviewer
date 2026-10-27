/**
 * PDF 양식(AcroForm) — 편집기에서 바꾼 값(방의 pdfForm: 칸 이름 → 값)을 저장할 때 채운다(설계 §2.3-4).
 * 값: 글자 칸·목록(하나)은 문자열, 체크 상자는 true/false, 라디오는 고른 값(내보내기 값), 여러 개 고르는 목록은 문자열 배열.
 * 한글 등 WinAnsi 밖 글자가 있으면 그 칸의 모양을 Pretendard로 다시 만든다(pdf-lib 기본 Helvetica로는 그릴 수 없어 저장이 실패한다).
 * 읽기 전용 칸·없는 칸·형식이 맞지 않는 값은 건너뛴다(저장 전체를 멈추지 않는다).
 */
import {
  PDFCheckBox,
  PDFDocument,
  PDFDropdown,
  PDFName,
  PDFOptionList,
  PDFRadioGroup,
  PDFTextField,
  type PDFFont,
} from 'pdf-lib';

export type PdfFormValue = string | boolean | readonly string[];

const WIN_ANSI = /^[\x20-\x7e\n\r\t]*$/;

export function isFormValue(value: unknown): value is PdfFormValue {
  if (typeof value === 'boolean') return true;
  if (typeof value === 'string') return value.length <= 20_000;
  return Array.isArray(value) && value.length <= 500 && value.every((v) => typeof v === 'string' && v.length <= 2000);
}

/** 문서에 양식이 있는지(없으면 getForm()이 빈 AcroForm을 만들어 버리므로 먼저 본다) */
export function hasAcroForm(doc: PDFDocument): boolean {
  return doc.catalog.has(PDFName.of('AcroForm'));
}

export async function applyFormValues(doc: PDFDocument, values: Iterable<[string, unknown]>, korean: () => Promise<PDFFont>): Promise<number> {
  if (!hasAcroForm(doc)) return 0;
  const form = doc.getForm();
  let applied = 0;
  for (const [name, value] of values) {
    if (!isFormValue(value)) continue;
    const field = form.getFieldMaybe(name);
    if (!field || field.isReadOnly()) continue;
    try {
      if (field instanceof PDFTextField && typeof value === 'string') {
        const max = field.getMaxLength();
        const text = max !== undefined && value.length > max ? value.slice(0, max) : value;
        field.setText(text);
        if (!WIN_ANSI.test(text)) field.updateAppearances(await korean());
      } else if (field instanceof PDFCheckBox && typeof value === 'boolean') {
        if (value) field.check();
        else field.uncheck();
      } else if (field instanceof PDFRadioGroup && typeof value === 'string') {
        // 편집기는 위젯의 켜짐 상태 이름을 보낸다. /Opt가 있는 칸은 상태 이름이 번호('0','1'…)이고 선택지 이름은 /Opt에 있다
        const options = field.getOptions();
        const option = options.includes(value) ? value : /^\d+$/.test(value) ? options[Number(value)] : undefined;
        if (value === '' || value === 'Off') field.clear();
        else if (option !== undefined) field.select(option);
        else continue;
      } else if (field instanceof PDFDropdown && typeof value === 'string') {
        if (value === '') field.clear();
        else field.select(value, field.isEditable());
        if (!WIN_ANSI.test(value)) field.updateAppearances(await korean());
      } else if (field instanceof PDFOptionList) {
        const list = (Array.isArray(value) ? value : typeof value === 'string' && value ? [value] : []).filter((v) => field.getOptions().includes(v));
        if (list.length === 0) field.clear();
        else field.select(list);
        if (list.some((v) => !WIN_ANSI.test(v))) field.updateAppearances(await korean());
      } else {
        continue;
      }
      applied += 1;
    } catch (error) {
      console.error(`양식 칸 "${name}" 채우기 실패(건너뜀):`, error);
    }
  }
  return applied;
}

/**
 * 지운 쪽에 있던 양식 칸 정리 — 그 쪽에만 있던 칸은 양식에서 빼고, 남은 위젯의 /P(쪽 가리킴)는 지운다.
 * 그래야 지운 쪽이 양식을 통해 파일에 남지 않는다(pdfModel.ts의 도달 불가 객체 정리).
 */
export function dropFieldsOnRemovedPages(doc: PDFDocument, removedAnnotRefs: ReadonlySet<string>): void {
  if (!hasAcroForm(doc) || removedAnnotRefs.size === 0) return;
  const form = doc.getForm();
  for (const field of form.getFields()) {
    const widgets = field.acroField.getWidgets();
    const refs = widgets.map((w) => doc.context.getObjectRef(w.dict)?.tag ?? (widgets.length === 1 ? field.ref.tag : ''));
    const removed = refs.map((tag) => removedAnnotRefs.has(tag));
    for (const [i, w] of widgets.entries()) if (removed[i]) w.dict.delete(PDFName.of('P'));
    if (widgets.length > 0 && removed.every(Boolean)) {
      try {
        form.acroForm.removeField(field.acroField);
      } catch {
        // 양식 구조가 특이하면 그대로 둔다(/P는 이미 끊었다)
      }
    }
  }
}
