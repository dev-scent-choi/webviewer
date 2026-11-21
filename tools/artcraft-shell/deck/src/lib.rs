include!("../../abi.rs");
use deckcraft_fonts::FontDb;
use deckcraft_model::Presentation;

struct Doc {
    doc: Presentation,
}

fn open_doc(name: &str, bytes: &[u8]) -> Result<Doc, String> {
    deckcraft_engine::cmd::file::open_presentation(name, bytes).map(|doc| Doc { doc }).map_err(|e| e.to_string())
}

fn count(d: &Doc) -> usize {
    d.doc.slides.len()
}

fn page_size(d: &Doc, i: usize) -> (f32, f32) {
    if i < d.doc.slides.len() { (d.doc.slide_size.width as f32, d.doc.slide_size.height as f32) } else { (0.0, 0.0) }
}

fn render(d: &Doc, i: usize, scale: f32) -> Result<(Vec<u8>, u32, u32), String> {
    if i >= d.doc.slides.len() {
        return Err(format!("no slide {i}"));
    }
    let scale = deckcraft_engine::cmd::file::image_scale(&d.doc, scale as f64);
    let opts = deckcraft_render::RenderOpts { scale, ..Default::default() };
    let img = deckcraft_render::render_slide(&d.doc, i, &opts);
    Ok((img.to_straight(), img.width, img.height))
}

// shortcut: deckcraft's FontDb has no alias API, so `alias` is ignored (Hangul relies on fallback).
fn add_font(bytes: Vec<u8>, _alias: &str) -> usize {
    FontDb::global().add_font(bytes)
}

fn text(d: &Doc, i: usize) -> Result<String, String> {
    if i >= d.doc.slides.len() {
        return Err(format!("no slide {i}"));
    }
    let mut runs = Vec::new();
    for pt in deckcraft_render::place_slide(&d.doc, i).texts {
        let mut chars: std::collections::HashMap<usize, Vec<char>> = Default::default();
        for run in &pt.layout.runs {
            if run.glyphs.is_empty() || run.size <= 0.0 {
                continue;
            }
            let cs = chars.entry(run.para).or_insert_with(|| {
                pt.body.paragraphs.get(run.para).map(|p| p.runs.iter().flat_map(|r| r.content().chars()).collect()).unwrap_or_default()
            });
            let (a, b) = (run.chars.0.min(cs.len()), run.chars.1.min(cs.len()));
            let t: String = cs.get(a..b.max(a)).unwrap_or_default().iter().collect();
            if t.trim().is_empty() {
                continue;
            }
            let (g0, gl) = (run.glyphs[0], run.glyphs[run.glyphs.len() - 1]);
            let end = gl.1 + run.face.advance(gl.0) * run.size / run.face.upem.max(1.0);
            // shortcut: rotated text gets its bounding box's left/right only.
            let p0 = pt.transform * kurbo::Point::new(g0.1, g0.2);
            let p1 = pt.transform * kurbo::Point::new(end, gl.2);
            runs.push(run_json(&t, p0.x.min(p1.x), p0.x.max(p1.x), p0.y, run.size));
        }
    }
    Ok(format!("{{\"w\":{},\"h\":{},\"runs\":[{}]}}", d.doc.slide_size.width, d.doc.slide_size.height, runs.join(",")))
}
