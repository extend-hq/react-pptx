use std::io::{Cursor, Write};

use pptx_core::{parse_presentation, ParseLimits, PresentationFormat, SlideNode};

fn record(version: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(payload.len() + 8);
    bytes.extend_from_slice(&version.to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

fn minimal_ppt() -> Vec<u8> {
    let utf16: Vec<u8> = "Legacy slide"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let text = record(0, 4000, &utf16);
    let slide = record(0x000f, 1006, &text);
    let cursor = Cursor::new(Vec::new());
    let mut compound = cfb::CompoundFile::create(cursor).unwrap();
    compound
        .create_stream("/PowerPoint Document")
        .unwrap()
        .write_all(&slide)
        .unwrap();
    compound.into_inner().into_inner()
}

#[test]
fn extracts_legacy_slide_text_into_the_shared_model() {
    let document = parse_presentation(&minimal_ppt(), &ParseLimits::default()).unwrap();
    assert_eq!(document.format, PresentationFormat::Ppt);
    assert_eq!(document.slides.len(), 1);
    let SlideNode::Shape { paragraphs, .. } = &document.slides[0].nodes[0] else {
        panic!("expected normalized text shape")
    };
    assert_eq!(paragraphs[0].runs[0].text, "Legacy slide");
    assert_eq!(document.warnings[0].code, "degraded-rendering");
}
