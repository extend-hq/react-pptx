use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum PresentationFormat {
    Pptx,
    Ppt,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationSize {
    pub width_emu: u64,
    pub height_emu: u64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationWarning {
    pub code: String,
    pub message: String,
    pub severity: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub slide_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub node_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub part_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub feature: Option<String>,
}

impl PresentationWarning {
    pub fn warning(code: &str, message: impl Into<String>) -> Self {
        Self {
            code: code.into(),
            message: message.into(),
            severity: "warning".into(),
            slide_index: None,
            node_id: None,
            part_name: None,
            feature: None,
        }
    }
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationMetadata {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subject: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub creator: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub company: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationAsset {
    pub id: String,
    pub content_type: String,
    pub byte_length: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data: Option<serde_bytes::ByteBuf>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationEmbeddedFont {
    pub family: String,
    pub asset_id: String,
    pub style: String,
    pub weight: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Transform {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rotation: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub flip_horizontal: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub flip_vertical: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase"
)]
#[allow(clippy::large_enum_variant)]
pub enum SlideNode {
    Shape {
        id: String,
        name: String,
        transform: Transform,
        geometry: ShapeGeometry,
        #[serde(skip_serializing_if = "Option::is_none")]
        fill: Option<FillStyle>,
        #[serde(skip_serializing_if = "Option::is_none")]
        line: Option<LineStyle>,
        paragraphs: Vec<TextParagraph>,
        #[serde(skip_serializing_if = "Option::is_none")]
        vertical_alignment: Option<VerticalAlignment>,
        #[serde(skip_serializing_if = "Option::is_none")]
        text_insets: Option<TextInsets>,
        #[serde(skip_serializing_if = "Option::is_none")]
        autofit: Option<TextAutofit>,
        #[serde(skip_serializing_if = "Option::is_none")]
        text_rotation: Option<f64>,
        #[serde(skip_serializing_if = "Option::is_none")]
        vertical_text: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        horizontal_overflow: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        vertical_overflow: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        text_wrap: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        column_count: Option<usize>,
        #[serde(skip_serializing_if = "Option::is_none")]
        column_spacing: Option<i64>,
        #[serde(skip_serializing_if = "Option::is_none")]
        right_to_left_columns: Option<bool>,
        #[serde(skip_serializing_if = "Option::is_none")]
        space_first_last_paragraph: Option<bool>,
    },
    Image {
        id: String,
        name: String,
        transform: Transform,
        asset_id: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        crop: Option<ImageCrop>,
        #[serde(skip_serializing_if = "Option::is_none")]
        opacity: Option<f64>,
        preserve_aspect_ratio: bool,
        #[serde(skip_serializing_if = "Option::is_none")]
        effects: Option<ImageEffects>,
    },
    Group {
        id: String,
        name: String,
        transform: Transform,
        children: Vec<SlideNode>,
        #[serde(skip_serializing_if = "Option::is_none")]
        child_transform: Option<Transform>,
    },
    Table {
        id: String,
        name: String,
        transform: Transform,
        rows: Vec<Vec<TableCell>>,
        #[serde(skip_serializing_if = "Vec::is_empty")]
        column_widths: Vec<i64>,
        #[serde(skip_serializing_if = "Vec::is_empty")]
        row_heights: Vec<i64>,
    },
    Chart {
        id: String,
        name: String,
        transform: Transform,
        chart_type: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        title: Option<String>,
        series: Vec<ChartSeries>,
        #[serde(skip_serializing_if = "Option::is_none")]
        has_legend: Option<bool>,
        /// Raw DrawingML chart part XML (`c:chartSpace` or `cx:chartSpace`) so
        /// viewers can render with full PowerPoint fidelity.
        #[serde(skip_serializing_if = "Option::is_none")]
        chart_xml: Option<String>,
        /// Companion Microsoft chart style part (`style*.xml`), when present.
        #[serde(skip_serializing_if = "Option::is_none")]
        chart_style_xml: Option<String>,
        /// Companion Microsoft chart color style part (`colors*.xml`), when present.
        #[serde(skip_serializing_if = "Option::is_none")]
        chart_colors_xml: Option<String>,
    },
    Unknown {
        id: String,
        name: String,
        transform: Transform,
        feature: String,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCell {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub row_span: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub col_span: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<FillStyle>,
    #[serde(skip_serializing_if = "BTreeMap::is_empty")]
    pub borders: BTreeMap<String, LineStyle>,
    pub paragraphs: Vec<TextParagraph>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_insets: Option<TextInsets>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical_alignment: Option<VerticalAlignment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<f64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextInsets {
    pub top: i64,
    pub right: i64,
    pub bottom: i64,
    pub left: i64,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageCrop {
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
    pub left: f64,
}

/// DrawingML picture color effects applied to an `a:blip`.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageEffects {
    /// `a:biLevel` black/white threshold, normalized to 0..1.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bi_level_threshold: Option<f64>,
    /// `a:grayscl` grayscale recolor.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub grayscale: Option<bool>,
    /// `a:duotone` two-color ramp, dark then light.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub duotone: Option<Vec<ColorValue>>,
    /// `a:lum` brightness, normalized to -1..1.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub brightness: Option<f64>,
    /// `a:lum` contrast, normalized to -1..1.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub contrast: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub categories: Option<Vec<serde_json::Value>>,
    pub values: Vec<Option<f64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ShapeGeometry {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub preset: Option<String>,
    /// SVG path data normalized to a 0..1 view box for DrawingML custom geometry.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ColorValue {
    pub value: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alpha: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GradientStop {
    pub position: f64,
    pub color: ColorValue,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum FillImageMode {
    Stretch,
    Tile,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "lowercase",
    rename_all_fields = "camelCase"
)]
pub enum FillStyle {
    None,
    Solid {
        color: ColorValue,
    },
    Gradient {
        #[serde(skip_serializing_if = "Option::is_none")]
        angle: Option<f64>,
        stops: Vec<GradientStop>,
    },
    Pattern {
        preset: String,
        foreground: ColorValue,
        background: ColorValue,
    },
    Image {
        asset_id: String,
        mode: FillImageMode,
        #[serde(skip_serializing_if = "Option::is_none")]
        crop: Option<ImageCrop>,
        #[serde(skip_serializing_if = "Option::is_none")]
        opacity: Option<f64>,
        #[serde(skip_serializing_if = "Option::is_none")]
        effects: Option<ImageEffects>,
    },
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LineStyle {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_arrow: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub end_arrow: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum VerticalAlignment {
    Top,
    Middle,
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum TextAutoFitMode {
    None,
    Normal,
    Shape,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextAutofit {
    pub mode: TextAutoFitMode,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_scale: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_spacing_reduction: Option<f64>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextParagraph {
    pub runs: Vec<TextRun>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alignment: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub level: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bullet: Option<TextBullet>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_spacing: Option<TextSpacing>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub space_before: Option<TextSpacing>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub space_after: Option<TextSpacing>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtl: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub margin_left_emu: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent_emu: Option<i64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextSpacing {
    pub value: f64,
    pub unit: TextSpacingUnit,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum TextSpacingUnit {
    Points,
    Percent,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBullet {
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_pt: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size_percent: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub start_at: Option<usize>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextRun {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_pt: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub baseline: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub language: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub character_spacing_pt: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationTheme {
    pub id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    pub colors: BTreeMap<String, String>,
    #[serde(skip_serializing_if = "BTreeMap::is_empty")]
    pub major_fonts: BTreeMap<String, String>,
    #[serde(skip_serializing_if = "BTreeMap::is_empty")]
    pub minor_fonts: BTreeMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideMaster {
    pub id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_id: Option<String>,
    pub nodes: Vec<SlideNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideLayout {
    pub id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub master_id: Option<String>,
    pub nodes: Vec<SlideNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationSlide {
    pub id: String,
    pub index: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub master_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub background: Option<FillStyle>,
    pub nodes: Vec<SlideNode>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub notes: Vec<SlideNote>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<SlideComment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_part: Option<String>,
    pub warnings: Vec<PresentationWarning>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideNote {
    pub text: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideComment {
    pub id: String,
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created_at: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub y: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationDocument {
    pub format: PresentationFormat,
    pub size: PresentationSize,
    pub slides: Vec<PresentationSlide>,
    pub masters: Vec<SlideMaster>,
    pub layouts: Vec<SlideLayout>,
    pub themes: Vec<PresentationTheme>,
    pub assets: BTreeMap<String, PresentationAsset>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub embedded_fonts: Vec<PresentationEmbeddedFont>,
    pub warnings: Vec<PresentationWarning>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub metadata: Option<PresentationMetadata>,
}
