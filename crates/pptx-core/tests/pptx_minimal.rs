use std::io::{Cursor, Write};

use pptx_core::{parse_presentation, ParseLimits, PresentationFormat, SlideNode};
use zip::{write::SimpleFileOptions, ZipWriter};

fn minimal_pptx() -> Vec<u8> {
    minimal_pptx_with_font(false)
}

fn minimal_pptx_with_font(include_font: bool) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut output);
    let options = SimpleFileOptions::default();
    zip.start_file("ppt/presentation.xml", options).unwrap();
    zip.write_all(if include_font {
        br#"<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:embeddedFontLst><p:embeddedFont><p:font typeface="Fixture Sans"/><p:regular r:id="rIdFont"/></p:embeddedFont></p:embeddedFontLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#.as_slice()
    } else {
        br#"<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#.as_slice()
    }).unwrap();
    zip.start_file("ppt/_rels/presentation.xml.rels", options)
        .unwrap();
    zip.write_all(if include_font {
        br#"<Relationships><Relationship Id="rId2" Target="slides/slide1.xml"/><Relationship Id="rIdFont" Target="fonts/font1.fntdata"/></Relationships>"#.as_slice()
    } else {
        br#"<Relationships><Relationship Id="rId2" Target="slides/slide1.xml"/></Relationships>"#.as_slice()
    }).unwrap();
    zip.start_file("ppt/slides/slide1.xml", options).unwrap();
    zip.write_all(br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="7" name="Title"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#).unwrap();
    if include_font {
        let mut font = vec![0u8; 48];
        font[..4].copy_from_slice(b"OTTO");
        zip.start_file("ppt/fonts/font1.fntdata", options).unwrap();
        zip.write_all(&font).unwrap();
    }
    zip.finish().unwrap();
    output.into_inner()
}

#[test]
fn extracts_embedded_powerpoint_fonts() {
    let document =
        parse_presentation(&minimal_pptx_with_font(true), &ParseLimits::default()).unwrap();
    assert_eq!(document.embedded_fonts.len(), 1);
    let font = &document.embedded_fonts[0];
    assert_eq!(font.family, "Fixture Sans");
    assert_eq!(font.style, "normal");
    assert_eq!(font.weight, "400");
    let asset = document.assets.get(&font.asset_id).unwrap();
    assert_eq!(asset.content_type, "font/otf");
    assert_eq!(&asset.data.as_ref().unwrap()[..4], b"OTTO");
}

#[test]
fn parses_slide_order_size_transform_and_text() {
    let document = parse_presentation(&minimal_pptx(), &ParseLimits::default()).unwrap();
    assert_eq!(document.format, PresentationFormat::Pptx);
    assert_eq!(document.size.width_emu, 12_192_000);
    assert_eq!(document.slides.len(), 1);
    let SlideNode::Shape {
        id,
        transform,
        paragraphs,
        ..
    } = &document.slides[0].nodes[0]
    else {
        panic!("expected shape")
    };
    assert_eq!(id, "7");
    assert_eq!(
        (transform.x, transform.y, transform.width, transform.height),
        (10, 20, 300, 400)
    );
    assert_eq!(paragraphs[0].runs[0].text, "Hello");
}

#[test]
fn enforces_input_limit_before_decompression() {
    let bytes = minimal_pptx();
    let limits = ParseLimits {
        max_input_bytes: 4,
        ..ParseLimits::default()
    };
    assert!(parse_presentation(&bytes, &limits).is_err());
}
