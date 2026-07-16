use std::io::{Cursor, Seek, Write};

use pptx_core::{ParseLimits, SlideNode, TextTabAlignment};
use zip::{write::SimpleFileOptions, ZipWriter};

fn add_file<W: Write + Seek>(zip: &mut ZipWriter<W>, path: &str, contents: &[u8]) {
    zip.start_file(path, SimpleFileOptions::default()).unwrap();
    zip.write_all(contents).unwrap();
}

fn text_layout_pptx() -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut output);
    add_file(
        &mut zip,
        "ppt/presentation.xml",
        br#"<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#,
    );
    add_file(
        &mut zip,
        "ppt/_rels/presentation.xml.rels",
        br#"<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/slide1.xml",
        r#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
          <p:sp><p:nvSpPr><p:cNvPr id="2" name="Text layout"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="3000000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
              <a:pPr algn="dist" marL="400050" indent="-228600" defTabSz="914400" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:tabLst><a:tab pos="1234440" algn="l"/><a:tab pos="2743200" algn="ctr"/><a:tab pos="4297680" algn="r"/><a:tab pos="5029200" algn="dec"/></a:tabLst>
              </a:pPr>
              <a:r><a:rPr lang="ja-JP" altLang="en-US" kern="1200"><a:latin typeface="Aptos"/><a:ea typeface="Yu Gothic"/><a:cs typeface="Arial"/><a:sym typeface="Symbol"/></a:rPr><a:t>日本語</a:t></a:r>
              <a:tab/>
              <a:r><a:rPr lang="ar-SA"><a:cs typeface="Noto Naskh Arabic"/><a:rtl/></a:rPr><a:t>العربية</a:t></a:r>
            </a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#
            .as_bytes(),
    );
    zip.finish().unwrap();
    output.into_inner()
}

#[test]
fn preserves_tabs_script_fonts_kerning_and_line_break_controls() {
    let document =
        pptx_core::parse_presentation(&text_layout_pptx(), &ParseLimits::default()).unwrap();
    let SlideNode::Shape { paragraphs, .. } = &document.slides[0].nodes[0] else {
        panic!("expected shape")
    };
    let paragraph = &paragraphs[0];
    assert_eq!(paragraph.alignment.as_deref(), Some("distributed"));
    assert_eq!(paragraph.margin_left_emu, Some(400_050));
    assert_eq!(paragraph.indent_emu, Some(-228_600));
    assert_eq!(paragraph.default_tab_size_emu, Some(914_400));
    assert_eq!(paragraph.east_asian_line_break, Some(true));
    assert_eq!(paragraph.latin_line_break, Some(false));
    assert_eq!(paragraph.hanging_punctuation, Some(true));
    let stops = paragraph.tab_stops.as_ref().unwrap();
    assert_eq!(stops.len(), 4);
    assert_eq!(stops[0].position_emu, 1_234_440);
    assert_eq!(stops[0].alignment, TextTabAlignment::Left);
    assert_eq!(stops[1].alignment, TextTabAlignment::Center);
    assert_eq!(stops[2].alignment, TextTabAlignment::Right);
    assert_eq!(stops[3].alignment, TextTabAlignment::Decimal);

    assert_eq!(paragraph.runs.len(), 3);
    let japanese = &paragraph.runs[0];
    assert_eq!(japanese.font_family.as_deref(), Some("Aptos"));
    assert_eq!(
        japanese.east_asian_font_family.as_deref(),
        Some("Yu Gothic")
    );
    assert_eq!(
        japanese.complex_script_font_family.as_deref(),
        Some("Arial")
    );
    assert_eq!(japanese.symbol_font_family.as_deref(), Some("Symbol"));
    assert_eq!(japanese.language.as_deref(), Some("ja-JP"));
    assert_eq!(japanese.alternative_language.as_deref(), Some("en-US"));
    assert_eq!(japanese.kerning_threshold_pt, Some(12.0));
    assert_eq!(paragraph.runs[1].text, "\t");
    let arabic = &paragraph.runs[2];
    assert_eq!(
        arabic.complex_script_font_family.as_deref(),
        Some("Noto Naskh Arabic")
    );
    assert_eq!(arabic.right_to_left, Some(true));
}
