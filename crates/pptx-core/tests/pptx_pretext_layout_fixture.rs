use pptx_core::{ParseLimits, SlideNode, TextAutoFitMode, TextTabAlignment};

const FIXTURE: &[u8] = include_bytes!("../../../tests/fixtures/pretext-layout-cases.pptx");

fn shape_named<'a>(nodes: &'a [SlideNode], name: &str) -> &'a SlideNode {
    nodes
        .iter()
        .find(|node| matches!(node, SlideNode::Shape { name: candidate, .. } if candidate == name))
        .unwrap_or_else(|| panic!("missing shape {name}"))
}

#[test]
fn parses_the_pretext_layout_regression_fixture() {
    let document = pptx_core::parse_presentation(FIXTURE, &ParseLimits::default()).unwrap();
    assert_eq!(document.slides.len(), 8);

    let SlideNode::Shape { paragraphs, .. } =
        shape_named(&document.slides[1].nodes, "JustifiedText")
    else {
        unreachable!()
    };
    assert_eq!(paragraphs[0].alignment.as_deref(), Some("justify"));
    let SlideNode::Shape { paragraphs, .. } =
        shape_named(&document.slides[1].nodes, "DistributedText")
    else {
        unreachable!()
    };
    assert_eq!(paragraphs[0].alignment.as_deref(), Some("distributed"));

    let SlideNode::Shape { paragraphs, .. } = shape_named(&document.slides[2].nodes, "SplitRuns")
    else {
        unreachable!()
    };
    assert!(paragraphs[0].runs.len() > 2);
    assert!(paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.runs)
        .all(|run| run.kerning_threshold_pt == Some(12.0)));
    assert!(paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.runs)
        .all(|run| run.character_spacing_pt == Some(0.75)));

    let SlideNode::Shape { paragraphs, .. } =
        shape_named(&document.slides[3].nodes, "ExplicitTabs")
    else {
        unreachable!()
    };
    assert!(paragraphs.iter().all(|paragraph| {
        paragraph
            .tab_stops
            .as_ref()
            .is_some_and(|stops| stops.len() == 3)
    }));
    assert!(paragraphs
        .iter()
        .all(|paragraph| paragraph.default_tab_size_emu == Some(914_400)));
    let stops = paragraphs[0].tab_stops.as_ref().unwrap();
    assert_eq!(stops[0].position_emu, 1_234_440);
    assert_eq!(stops[0].alignment, TextTabAlignment::Left);
    assert!(paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.runs)
        .any(|run| run.text.contains('\t')));

    let SlideNode::Shape { paragraphs, .. } =
        shape_named(&document.slides[3].nodes, "HangingIndent")
    else {
        unreachable!()
    };
    assert_eq!(paragraphs[0].margin_left_emu, Some(400_050));
    assert_eq!(paragraphs[0].indent_emu, Some(-228_600));

    let SlideNode::Shape { paragraphs, .. } =
        shape_named(&document.slides[4].nodes, "JapaneseText")
    else {
        unreachable!()
    };
    assert_eq!(paragraphs[0].east_asian_line_break, Some(true));
    assert_eq!(paragraphs[0].hanging_punctuation, Some(true));
    assert_eq!(paragraphs[0].runs[0].language.as_deref(), Some("ja-JP"));
    assert_eq!(
        paragraphs[0].runs[0].east_asian_font_family.as_deref(),
        Some("Yu Gothic")
    );

    let SlideNode::Shape { paragraphs, .. } = shape_named(&document.slides[4].nodes, "ArabicText")
    else {
        unreachable!()
    };
    assert_eq!(paragraphs[0].rtl, Some(true));
    assert_eq!(paragraphs[0].runs[0].language.as_deref(), Some("ar-SA"));
    assert_eq!(paragraphs[0].runs[0].right_to_left, Some(true));
    assert_eq!(
        paragraphs[0].runs[0].complex_script_font_family.as_deref(),
        Some("Noto Naskh Arabic")
    );

    let SlideNode::Shape {
        column_count,
        column_spacing,
        ..
    } = shape_named(&document.slides[5].nodes, "TwoColumns")
    else {
        unreachable!()
    };
    assert_eq!(*column_count, Some(2));
    assert_eq!(*column_spacing, Some(182_880));
    let SlideNode::Shape { autofit, .. } = shape_named(&document.slides[5].nodes, "NormalAutofit")
    else {
        unreachable!()
    };
    assert_eq!(autofit.unwrap().mode, TextAutoFitMode::Normal);
    assert_eq!(autofit.unwrap().font_scale, Some(0.88));
    assert_eq!(autofit.unwrap().line_spacing_reduction, Some(0.12));
}
