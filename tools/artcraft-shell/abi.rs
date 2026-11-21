// ArtCraft 브라우저 엔진 공통 C ABI — 엔진 껍데기(word·deck)가 include!로 넣는다.
// JS: p = sp_alloc(n); 바이트 복사; 호출; sp_out_ptr/len, sp_out_w/h 읽기; sp_free(p, n).
use std::cell::RefCell;

thread_local! {
    static OUT: RefCell<(Vec<u8>, u32, u32)> = RefCell::new((Vec::new(), 0, 0));
    static MSG: RefCell<String> = RefCell::new(String::new());
    static DOCS: RefCell<Vec<Option<Doc>>> = RefCell::new(Vec::new());
}

fn layout(n: usize) -> Option<std::alloc::Layout> {
    std::alloc::Layout::array::<u8>(n.max(1)).ok()
}

/// n바이트 버퍼. 너무 크거나 메모리가 없으면 0(null) — JS는 0이면 실패로 본다.
#[unsafe(no_mangle)]
pub extern "C" fn sp_alloc(n: usize) -> *mut u8 {
    layout(n).map_or(std::ptr::null_mut(), |l| unsafe { std::alloc::alloc(l) })
}

/// sp_alloc(n)이 준 p를 같은 n으로 돌려준다. null은 무시.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn sp_free(p: *mut u8, n: usize) {
    if let (false, Some(l)) = (p.is_null(), layout(n)) {
        unsafe { std::alloc::dealloc(p, l) }
    }
}

/// 길이 0이거나 null이면 빈 조각 — from_raw_parts는 n == 0이어도 null이면 UB다(빈 alias를 (0, 0)으로 넘기는 경우).
unsafe fn slice<'a>(p: *const u8, n: usize) -> &'a [u8] {
    if n == 0 || p.is_null() {
        return &[];
    }
    unsafe { std::slice::from_raw_parts(p, n) }
}

fn set_out(v: Vec<u8>, w: u32, h: u32) {
    OUT.with(|o| *o.borrow_mut() = (v, w, h));
}
fn set_msg(s: String) {
    MSG.with(|m| *m.borrow_mut() = s);
}

/// 핸들 h(1부터)의 문서로 f를 부른다. 없는 핸들이면 None.
fn with_doc<R>(h: u32, f: impl FnOnce(&Doc) -> R) -> Option<R> {
    DOCS.with(|ds| ds.borrow().get((h as usize).wrapping_sub(1)).and_then(|d| d.as_ref()).map(f))
}

#[unsafe(no_mangle)]
pub extern "C" fn sp_out_ptr() -> *const u8 { OUT.with(|o| o.borrow().0.as_ptr()) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_out_len() -> usize { OUT.with(|o| o.borrow().0.len()) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_out_w() -> u32 { OUT.with(|o| o.borrow().1) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_out_h() -> u32 { OUT.with(|o| o.borrow().2) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_out_clear() { set_out(Vec::new(), 0, 0) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_msg_ptr() -> *const u8 { MSG.with(|m| m.borrow().as_ptr()) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_msg_len() -> usize { MSG.with(|m| m.borrow().len()) }

/// 이름의 확장자로 형식을 고른다. 핸들(1 이상), 실패면 0(sp_msg에 이유).
#[unsafe(no_mangle)]
pub unsafe extern "C" fn sp_open(name: *const u8, name_len: usize, p: *const u8, n: usize) -> u32 {
    let name = String::from_utf8_lossy(unsafe { slice(name, name_len) }).into_owned();
    match open_doc(&name, unsafe { slice(p, n) }) {
        Ok(d) => DOCS.with(|ds| {
            let mut ds = ds.borrow_mut();
            ds.push(Some(d));
            ds.len() as u32
        }),
        Err(e) => {
            set_msg(e);
            0
        }
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn sp_count(h: u32) -> u32 {
    with_doc(h, |d| count(d) as u32).unwrap_or(0)
}

/// i쪽(0부터)의 폭(pt). 없으면 0.
#[unsafe(no_mangle)]
pub extern "C" fn sp_page_w(h: u32, i: u32) -> f32 {
    with_doc(h, |d| page_size(d, i as usize).0).unwrap_or(0.0)
}

/// i쪽(0부터)의 높이(pt). 없으면 0.
#[unsafe(no_mangle)]
pub extern "C" fn sp_page_h(h: u32, i: u32) -> f32 {
    with_doc(h, |d| page_size(d, i as usize).1).unwrap_or(0.0)
}

/// i쪽을 scale_milli/1000 px/pt로 OUT에(RGBA8 직선 알파). 0 성공, 1 실패.
#[unsafe(no_mangle)]
pub extern "C" fn sp_render(h: u32, i: u32, scale_milli: u32) -> u32 {
    match with_doc(h, |d| render(d, i as usize, scale_milli as f32 / 1000.0)).unwrap_or_else(|| Err("bad handle".into())) {
        Ok((v, w, hh)) => {
            set_out(v, w, hh);
            0
        }
        Err(e) => {
            set_msg(e);
            1
        }
    }
}

/// i쪽의 글자 조각 JSON을 OUT에: {"w":pt,"h":pt,"runs":[{"t","x0","x1","y"(기준선),"s"}]}, pt, y 아래로. 0 성공, 1 실패.
#[unsafe(no_mangle)]
pub extern "C" fn sp_text(h: u32, i: u32) -> u32 {
    match with_doc(h, |d| text(d, i as usize)).unwrap_or_else(|| Err("bad handle".into())) {
        Ok(s) => {
            set_out(s.into_bytes(), 0, 0);
            0
        }
        Err(e) => {
            set_msg(e);
            1
        }
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn sp_close(h: u32) {
    DOCS.with(|ds| {
        if let Some(d) = ds.borrow_mut().get_mut((h as usize).wrapping_sub(1)) {
            *d = None;
        }
    });
}

/// 글꼴(TTF/OTF) 넣기. alias(빈 값 가능)는 엔진이 허락하면 그 이름으로도 등록(메모리 주의 — 화면은 빈 값만 쓴다). 넣은 face 수.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn sp_add_font(p: *const u8, n: usize, alias: *const u8, alias_len: usize) -> u32 {
    let alias = String::from_utf8_lossy(unsafe { slice(alias, alias_len) }).into_owned();
    add_font(unsafe { slice(p, n) }.to_vec(), &alias) as u32
}

fn json_str(s: &str) -> String {
    let mut o = String::with_capacity(s.len() + 2);
    o.push('"');
    for c in s.chars() {
        match c {
            '"' => o.push_str("\\\""),
            '\\' => o.push_str("\\\\"),
            c if (c as u32) < 0x20 => o.push_str(&format!("\\u{:04x}", c as u32)),
            c => o.push(c),
        }
    }
    o.push('"');
    o
}

fn run_json(t: &str, x0: f64, x1: f64, y: f64, s: f64) -> String {
    format!("{{\"t\":{},\"x0\":{:.1},\"x1\":{:.1},\"y\":{:.1},\"s\":{:.1}}}", json_str(t), x0, x1, y, s)
}
