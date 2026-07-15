use std::collections::BTreeMap;

use pptx_core::{
    ChartSeries, ColorValue, FillStyle, LineStyle, ShapeGeometry, SlideNode, TableCell, TextBullet,
    TextParagraph, TextRun, TextSpacing, TextSpacingUnit, Transform, VerticalAlignment,
};
use serde_json::json;

fn shape(
    fill: Option<FillStyle>,
    line: Option<LineStyle>,
    vertical_alignment: Option<VerticalAlignment>,
) -> SlideNode {
    SlideNode::Shape {
        id: "shape-1".into(),
        name: "Styled shape".into(),
        transform: Transform {
            x: 1,
            y: 2,
            width: 3,
            height: 4,
            ..Default::default()
        },
        geometry: ShapeGeometry {
            preset: Some("rect".into()),
            path: None,
        },
        fill,
        line,
        paragraphs: Vec::new(),
        vertical_alignment,
        text_insets: None,
        autofit: None,
        text_rotation: None,
        vertical_text: None,
        horizontal_overflow: None,
        vertical_overflow: None,
        text_wrap: None,
        column_count: None,
        column_spacing: None,
        right_to_left_columns: None,
        space_first_last_paragraph: None,
    }
}

#[test]
fn shape_styles_serialize_to_the_typescript_model_contract() {
    let node = shape(
        Some(FillStyle::Solid {
            color: ColorValue {
                value: "#112233".into(),
                alpha: Some(0.5),
            },
        }),
        Some(LineStyle {
            color: Some(ColorValue {
                value: "#445566".into(),
                alpha: None,
            }),
            width: Some(2.25),
            start_arrow: Some("triangle".into()),
            ..Default::default()
        }),
        Some(VerticalAlignment::Middle),
    );

    let value = serde_json::to_value(node).unwrap();
    assert_eq!(
        value["fill"],
        json!({
            "type": "solid",
            "color": { "value": "#112233", "alpha": 0.5 }
        })
    );
    assert_eq!(
        value["line"],
        json!({
            "color": { "value": "#445566" },
            "width": 2.25,
            "startArrow": "triangle"
        })
    );
    assert_eq!(value["verticalAlignment"], "middle");
    assert!(value.get("vertical_alignment").is_none());
}

#[test]
fn absent_shape_styles_are_omitted() {
    let value = serde_json::to_value(shape(None, None, None)).unwrap();
    let object = value.as_object().unwrap();

    assert!(object.get("fill").is_none());
    assert!(object.get("line").is_none());
    assert!(object.get("verticalAlignment").is_none());
}

#[test]
fn image_fields_serialize_to_the_typescript_model_contract() {
    let node = SlideNode::Image {
        id: "image-1".into(),
        name: "Picture".into(),
        transform: Transform {
            x: 1,
            y: 2,
            width: 3,
            height: 4,
            ..Default::default()
        },
        asset_id: "asset-1".into(),
        crop: None,
        opacity: None,
        preserve_aspect_ratio: true,
    };

    let value = serde_json::to_value(node).unwrap();
    assert_eq!(value["assetId"], "asset-1");
    assert_eq!(value["preserveAspectRatio"], true);
    assert!(value.get("asset_id").is_none());
    assert!(value.get("preserve_aspect_ratio").is_none());
}

#[test]
fn structured_nodes_serialize_to_the_typescript_model_contract() {
    let group = SlideNode::Group {
        id: "group-1".into(),
        name: "Group".into(),
        transform: Transform::default(),
        children: Vec::new(),
        child_transform: Some(Transform {
            x: 1,
            y: 2,
            width: 3,
            height: 4,
            ..Default::default()
        }),
    };
    let table = SlideNode::Table {
        id: "table-1".into(),
        name: "Table".into(),
        transform: Transform::default(),
        rows: vec![vec![TableCell {
            row_span: None,
            col_span: Some(2),
            fill: None,
            borders: BTreeMap::new(),
            paragraphs: vec![TextParagraph::default()],
            text_insets: None,
            vertical_alignment: None,
            text_rotation: None,
        }]],
        column_widths: vec![100, 200],
        row_heights: vec![50],
    };
    let chart = SlideNode::Chart {
        id: "chart-1".into(),
        name: "Chart".into(),
        transform: Transform::default(),
        chart_type: "bar".into(),
        title: Some("Revenue".into()),
        series: vec![ChartSeries {
            name: Some("Baseline".into()),
            categories: Some(vec![json!("Q1")]),
            values: vec![Some(42.0)],
            color: None,
        }],
        has_legend: Some(true),
    };

    let group = serde_json::to_value(group).unwrap();
    assert_eq!(group["type"], "group");
    assert_eq!(group["childTransform"]["width"], 3);
    assert!(group.get("child_transform").is_none());

    let table = serde_json::to_value(table).unwrap();
    assert_eq!(table["type"], "table");
    assert_eq!(table["rows"][0][0]["colSpan"], 2);
    assert_eq!(table["columnWidths"], json!([100, 200]));
    assert_eq!(table["rowHeights"], json!([50]));

    let chart = serde_json::to_value(chart).unwrap();
    assert_eq!(chart["type"], "chart");
    assert_eq!(chart["chartType"], "bar");
    assert_eq!(chart["series"][0]["categories"], json!(["Q1"]));
    assert_eq!(chart["series"][0]["values"], json!([42.0]));
    assert_eq!(chart["hasLegend"], true);
}

#[test]
fn rich_text_fields_serialize_to_the_typescript_model_contract() {
    let paragraph = TextParagraph {
        runs: vec![TextRun {
            text: "Linked".into(),
            font_family: Some("Aptos".into()),
            font_size_pt: Some(18.0),
            bold: Some(true),
            italic: Some(false),
            underline: Some(true),
            strike: Some(true),
            color: Some(ColorValue {
                value: "#123456".into(),
                alpha: None,
            }),
            baseline: Some(30.0),
            language: Some("en-US".into()),
            hyperlink: Some("https://example.com".into()),
            character_spacing_pt: Some(-0.08),
        }],
        alignment: Some("center".into()),
        level: Some(1),
        bullet: Some(TextBullet {
            kind: "character".into(),
            value: Some("*".into()),
            font_family: None,
            font_size_pt: None,
            size_percent: None,
            start_at: None,
        }),
        line_spacing: Some(TextSpacing {
            value: 1.2,
            unit: TextSpacingUnit::Percent,
        }),
        space_before: Some(TextSpacing {
            value: 6.0,
            unit: TextSpacingUnit::Points,
        }),
        space_after: Some(TextSpacing {
            value: 3.0,
            unit: TextSpacingUnit::Points,
        }),
        rtl: Some(true),
        margin_left_emu: None,
        indent_emu: None,
    };

    let value = serde_json::to_value(paragraph).unwrap();
    assert_eq!(value["alignment"], "center");
    assert_eq!(
        value["lineSpacing"],
        json!({"value": 1.2, "unit": "percent"})
    );
    assert_eq!(
        value["spaceBefore"],
        json!({"value": 6.0, "unit": "points"})
    );
    assert_eq!(value["spaceAfter"], json!({"value": 3.0, "unit": "points"}));
    assert_eq!(value["runs"][0]["fontFamily"], "Aptos");
    assert_eq!(value["runs"][0]["fontSizePt"], 18.0);
    assert_eq!(value["runs"][0]["color"]["value"], "#123456");
    assert_eq!(value["runs"][0]["hyperlink"], "https://example.com");
    assert_eq!(value["runs"][0]["characterSpacingPt"], -0.08);
    assert!(value["runs"][0].get("font_family").is_none());
}
