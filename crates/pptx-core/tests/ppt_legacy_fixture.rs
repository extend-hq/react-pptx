use pptx_core::{parse_presentation, FillStyle, ParseLimits, PresentationFormat, SlideNode};

const LEGACY_DECK: &[u8] = include_bytes!("../../../tests/fixtures/legacy/file-example-250kb.ppt");

fn document_text(document: &pptx_core::PresentationDocument) -> String {
    document
        .slides
        .iter()
        .flat_map(|slide| slide.nodes.iter())
        .filter_map(|node| match node {
            SlideNode::Shape { paragraphs, .. } => Some(paragraphs),
            _ => None,
        })
        .flat_map(|paragraphs| paragraphs.iter())
        .flat_map(|paragraph| paragraph.runs.iter())
        .map(|run| run.text.as_str())
        .collect::<Vec<_>>()
        .join("\n")
}

#[test]
fn parses_the_real_world_legacy_fixture_into_renderable_slides() {
    let document = parse_presentation(LEGACY_DECK, &ParseLimits::default())
        .expect("the supplied PowerPoint 97-2003 fixture should parse");

    assert_eq!(document.format, PresentationFormat::Ppt);
    assert_eq!(document.slides.len(), 3);
    assert_eq!(document.size.width_emu, 10_080_625);
    assert_eq!(document.size.height_emu, 7_559_675);
    assert!(
        document.slides.iter().all(|slide| !slide.nodes.is_empty()),
        "every source slide should produce renderable normalized nodes"
    );

    let text = document_text(&document);
    assert!(text.contains("Lorem ipsum"));
    assert!(text.contains("Table"));
    assert!(
        !text.contains("Click to edit the title text format"),
        "slide-master placeholder text must not become a visible slide"
    );
    assert!(
        document
            .warnings
            .iter()
            .any(|warning| warning.code == "degraded-rendering"),
        "unsupported binary-format features must be surfaced explicitly"
    );

    let png = document
        .assets
        .values()
        .find(|asset| asset.content_type == "image/png")
        .expect("the referenced master background PNG should be extracted");
    assert!(png
        .data
        .as_deref()
        .is_some_and(|data| data.starts_with(b"\x89PNG\r\n\x1a\n")));
    let emf = document
        .assets
        .values()
        .find(|asset| asset.content_type == "image/x-emf")
        .expect("the chart preview EMF should be decompressed and extracted");
    assert_eq!(emf.byte_length, 13_932);

    let slide_two_image_assets = document.slides[1]
        .nodes
        .iter()
        .filter_map(|node| match node {
            SlideNode::Image { asset_id, .. } => Some(asset_id.as_str()),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert!(slide_two_image_assets.iter().any(|id| *id == png.id));
    assert!(slide_two_image_assets.iter().any(|id| *id == emf.id));
    let chart_preview = document.slides[1]
        .nodes
        .iter()
        .find_map(|node| match node {
            SlideNode::Image {
                asset_id,
                transform,
                ..
            } if asset_id == &emf.id => Some(transform),
            _ => None,
        })
        .expect("the chart preview should retain its OfficeArt anchor");
    assert_eq!(chart_preview.x, 1_693_862);
    assert_eq!(chart_preview.y, 2_487_612);
    assert_eq!(chart_preview.width, 6_865_938);
    assert_eq!(chart_preview.height, 3_598_863);
    assert!(document.slides[0].nodes.iter().any(|node| matches!(
        node,
        SlideNode::Image {
            asset_id,
            preserve_aspect_ratio: false,
            ..
        } if asset_id == &png.id
    )));

    let title_run = document.slides[0]
        .nodes
        .iter()
        .find_map(|node| match node {
            SlideNode::Shape { paragraphs, .. } => paragraphs
                .iter()
                .flat_map(|paragraph| paragraph.runs.iter())
                .find(|run| run.text == "Lorem ipsum"),
            _ => None,
        })
        .expect("slide title should be associated with its OfficeArt text shape");
    assert_eq!(title_run.font_family.as_deref(), Some("Arial"));
    assert_eq!(title_run.font_size_pt, Some(36.0));

    let table_fills = document.slides[2]
        .nodes
        .iter()
        .filter_map(|node| match node {
            SlideNode::Shape {
                fill: Some(FillStyle::Solid { color }),
                ..
            } => Some(color.value.as_str()),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert!(
        table_fills
            .iter()
            .any(|color| matches!(*color, "#B3B3B3" | "#CCCCCC" | "#E6E6E6")),
        "direct OfficeArt fill colors should survive normalization"
    );
    assert!(
        document.slides[2].nodes.len() >= 50,
        "the table's OfficeArt cell and line geometry should be preserved"
    );
}
