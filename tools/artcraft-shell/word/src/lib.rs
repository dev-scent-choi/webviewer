include!("../../abi.rs");
use wordcraft_doc::Document;
use wordcraft_fonts::FontDb;

struct Doc {
    doc: Document,
    layout: wordcraft_layout::DocLayout,
}

fn open_doc(name: &str, bytes: &[u8]) -> Result<Doc, String> {
    let doc = wordcraft_engine::io::open_bytes(name, bytes)?;
    let doc = wordcraft_engine::cmd::references::output_fields(&doc).unwrap_or(doc);
    let lo = wordcraft_layout::LayoutOptions { output_page_refs: true, ..Default::default() };
    let layout = wordcraft_layout::layout(&doc, &mut wordcraft_layout::LayoutCache::new(), &lo);
    Ok(Doc { doc, layout })
}

fn count(d: &Doc) -> usize {
    d.layout.pages.len()
}

fn page_size(d: &Doc, i: usize) -> (f32, f32) {
    d.layout.pages.get(i).map_or((0.0, 0.0), |p| (p.w, p.h))
}

fn render(d: &Doc, i: usize, scale: f32) -> Result<(Vec<u8>, u32, u32), String> {
    let p = d.layout.pages.get(i).ok_or_else(|| format!("no page {}", i + 1))?;
    let img = wordcraft_render::render_page(&d.doc, p, scale, &Default::default());
    Ok((img.to_straight(), img.width, img.height))
}

fn add_font(bytes: Vec<u8>, alias: &str) -> usize {
    if alias.is_empty() { FontDb::global().add_font(bytes) } else { FontDb::global().add_document_fonts(vec![(bytes, alias.to_string())]) }
}

fn text(d: &Doc, i: usize) -> Result<String, String> {
    use wordcraft_layout::display::{Draw, page_display};
    let p = d.layout.pages.get(i).ok_or_else(|| format!("no page {}", i + 1))?;
    // shortcut: glyphs inside Turned/Rotated (vertical cell text, rotated shapes) are skipped; add their transform if search must cover them.
    fn walk(ds: &[Draw], out: &mut Vec<String>) {
        for dr in ds {
            match dr {
                Draw::Glyphs { face, size, glyphs, text, .. } if !glyphs.is_empty() && !text.trim().is_empty() => {
                    let (g0, gl) = (glyphs[0], glyphs[glyphs.len() - 1]);
                    let adv = face.advance(gl.0) * *size as f64 / face.units_per_em().max(1.0);
                    out.push(run_json(text, g0.1 as f64, gl.1 as f64 + adv, g0.2 as f64, *size as f64));
                }
                Draw::Figure { draws, .. } => walk(draws, out),
                Draw::Clip { items, .. } => walk(items, out),
                _ => {}
            }
        }
    }
    let mut runs = Vec::new();
    walk(&page_display(&d.doc, p, &Default::default()), &mut runs);
    Ok(format!("{{\"w\":{},\"h\":{},\"runs\":[{}]}}", p.w, p.h, runs.join(",")))
}
