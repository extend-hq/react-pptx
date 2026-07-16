use std::io::{Cursor, Seek, Write};

use pptx_core::{
    FillImageMode, FillStyle, ParseError, ParseLimits, SlideNode, TextAutoFitMode, TextSpacing,
    TextSpacingUnit, VerticalAlignment,
};
use zip::{write::SimpleFileOptions, ZipWriter};

fn add_file<W: Write + Seek>(zip: &mut ZipWriter<W>, path: &str, contents: &[u8]) {
    zip.start_file(path, SimpleFileOptions::default()).unwrap();
    zip.write_all(contents).unwrap();
}

fn feature_pptx() -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut output);

    add_file(
        &mut zip,
        "ppt/presentation.xml",
        br#"<p:presentation xmlns:p="p" xmlns:r="r">
          <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
          <p:sldSz cx="12192000" cy="6858000"/>
        </p:presentation>"#,
    );
    add_file(
        &mut zip,
        "ppt/_rels/presentation.xml.rels",
        br#"<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/_rels/slide1.xml.rels",
        br#"<Relationships>
          <Relationship Id="rIdImage" Target="../media/image1.png"/>
          <Relationship Id="rIdChart" Target="../charts/chart1.xml"/>
          <Relationship Id="rIdLink" Target="https://example.com/docs" TargetMode="External"/>
        </Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/slide1.xml",
        br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:c="c">
          <p:cSld>
            <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F4F0E8"/></a:solidFill></p:bgPr></p:bg>
            <p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
              <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>

              <p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Styled title"/></p:nvSpPr>
                <p:spPr>
                  <a:xfrm rot="1080000" flipH="1"><a:off x="100" y="200"/><a:ext cx="3000" cy="4000"/></a:xfrm>
                  <a:prstGeom prst="roundRect"/>
                  <a:gradFill>
                    <a:gsLst>
                      <a:gs pos="0"><a:srgbClr val="112233"/></a:gs>
                      <a:gs pos="100000"><a:srgbClr val="445566"><a:alpha val="50000"/></a:srgbClr></a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000"/>
                  </a:gradFill>
                  <a:ln w="38100">
                    <a:solidFill><a:srgbClr val="D8663A"/></a:solidFill>
                    <a:prstDash val="dash"/><a:headEnd type="triangle"/><a:tailEnd type="oval"/>
                  </a:ln>
                </p:spPr>
                <p:txBody>
                  <a:bodyPr anchor="b"/>
                  <a:p>
                    <a:pPr algn="ctr" lvl="1" rtl="1"><a:buChar char="*"/><a:lnSpc><a:spcPts val="1800"/></a:lnSpc></a:pPr>
                    <a:r><a:rPr lang="en-US" sz="2700" b="1" u="sng" strike="sngStrike" baseline="30000"><a:latin typeface="Aptos Display"/><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>First</a:t></a:r>
                    <a:r><a:rPr lang="en-US" sz="1600" i="1"><a:latin typeface="Aptos"/></a:rPr><a:t> run</a:t></a:r>
                  </a:p>
                  <a:p><a:r><a:rPr sz="1200"/><a:t>Second paragraph</a:t></a:r></a:p>
                </p:txBody>
              </p:sp>

              <p:pic>
                <p:nvPicPr><p:cNvPr id="3" name="Picture"/></p:nvPicPr>
                <p:blipFill><a:blip r:embed="rIdImage"><a:biLevel thresh="60000"/></a:blip><a:stretch/></p:blipFill>
                <p:spPr><a:xfrm flipV="true"><a:off x="5000" y="6000"/><a:ext cx="7000" cy="8000"/></a:xfrm></p:spPr>
              </p:pic>

              <p:grpSp>
                <p:nvGrpSpPr><p:cNvPr id="4" name="Group"/></p:nvGrpSpPr>
                <p:grpSpPr><a:xfrm><a:off x="10000" y="11000"/><a:ext cx="12000" cy="13000"/><a:chOff x="10" y="20"/><a:chExt cx="100" cy="200"/></a:xfrm></p:grpSpPr>
                <p:sp>
                  <p:nvSpPr><p:cNvPr id="5" name="Grouped shape"/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm><a:prstGeom prst="ellipse"/><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></p:spPr>
                </p:sp>
              </p:grpSp>

              <p:graphicFrame>
                <p:nvGraphicFramePr><p:cNvPr id="6" name="Table"/></p:nvGraphicFramePr>
                <p:xfrm><a:off x="20000" y="21000"/><a:ext cx="22000" cy="23000"/></p:xfrm>
                <a:graphic><a:graphicData>
                  <a:tbl>
                    <a:tblGrid><a:gridCol w="1000"/><a:gridCol w="2000"/></a:tblGrid>
                    <a:tr h="500">
                      <a:tc gridSpan="2">
                        <a:txBody><a:p><a:r><a:rPr sz="1400" b="1"><a:latin typeface="Aptos"/></a:rPr><a:t>Header</a:t></a:r></a:p></a:txBody>
                        <a:tcPr><a:solidFill><a:srgbClr val="14362C"/></a:solidFill><a:lnT w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnT></a:tcPr>
                      </a:tc>
                      <a:tc hMerge="1"><a:txBody><a:p><a:r><a:t>continuation</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                    </a:tr>
                    <a:tr h="600">
                      <a:tc><a:txBody><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                      <a:tc><a:txBody><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                    </a:tr>
                  </a:tbl>
                </a:graphicData></a:graphic>
              </p:graphicFrame>

              <p:graphicFrame>
                <p:nvGraphicFramePr><p:cNvPr id="7" name="Chart"/></p:nvGraphicFramePr>
                <p:xfrm><a:off x="30000" y="31000"/><a:ext cx="32000" cy="33000"/></p:xfrm>
                <a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
              </p:graphicFrame>
            </p:spTree>
          </p:cSld>
        </p:sld>"#,
    );
    add_file(
        &mut zip,
        "ppt/charts/chart1.xml",
        br#"<c:chartSpace xmlns:c="c" xmlns:a="a">
          <c:chart>
            <c:title><c:tx><c:rich><a:p><a:r><a:t>Performance</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:plotArea><c:barChart>
              <c:ser>
                <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Baseline</c:v></c:pt></c:strCache></c:strRef></c:tx>
                <c:spPr><a:solidFill><a:srgbClr val="D8663A"/></a:solidFill></c:spPr>
                <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>001</c:v></c:pt><c:pt idx="1"><c:v>Render</c:v></c:pt></c:strCache></c:strRef></c:cat>
                <c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>700</c:v></c:pt><c:pt idx="2"><c:v>280</c:v></c:pt></c:numCache></c:numRef></c:val>
              </c:ser>
              <c:ser>
                <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Budget</c:v></c:pt></c:strCache></c:strRef></c:tx>
                <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Parse</c:v></c:pt><c:pt idx="1"><c:v>Render</c:v></c:pt></c:strCache></c:strRef></c:cat>
                <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>900</c:v></c:pt><c:pt idx="1"><c:v>400</c:v></c:pt></c:numCache></c:numRef></c:val>
              </c:ser>
            </c:barChart></c:plotArea>
            <c:legend/>
          </c:chart>
        </c:chartSpace>"#,
    );
    add_file(
        &mut zip,
        "ppt/charts/_rels/chart1.xml.rels",
        br#"<Relationships>
          <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
        </Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/charts/style1.xml",
        br#"<cs:chartStyle xmlns:cs="cs"/>"#,
    );
    add_file(
        &mut zip,
        "ppt/charts/colors1.xml",
        br#"<cs:colorStyle xmlns:cs="cs"/>"#,
    );
    add_file(
        &mut zip,
        "ppt/media/image1.png",
        b"\x89PNG\r\n\x1a\nfixture",
    );

    zip.finish().unwrap();
    output.into_inner()
}

#[test]
fn parses_owned_presentationml_rendering_features() {
    let document = pptx_core::parse_presentation(&feature_pptx(), &ParseLimits::default()).unwrap();
    let slide = &document.slides[0];
    assert_eq!(slide.nodes.len(), 5);
    assert!(
        matches!(&slide.background, Some(FillStyle::Solid { color }) if color.value == "#F4F0E8")
    );

    let SlideNode::Shape {
        transform,
        fill,
        line,
        paragraphs,
        vertical_alignment,
        ..
    } = &slide.nodes[0]
    else {
        panic!("expected styled shape")
    };
    assert_eq!(transform.rotation, Some(18.0));
    assert_eq!(transform.flip_horizontal, Some(true));
    let Some(FillStyle::Gradient { angle, stops }) = fill else {
        panic!("expected gradient fill")
    };
    assert_eq!(*angle, Some(90.0));
    assert_eq!(stops.len(), 2);
    assert_eq!(stops[1].color.alpha, Some(0.5));
    let line = line.as_ref().unwrap();
    assert_eq!(line.width, Some(4.0));
    assert_eq!(line.dash.as_deref(), Some("dash"));
    assert_eq!(line.start_arrow.as_deref(), Some("triangle"));
    assert_eq!(line.end_arrow.as_deref(), Some("oval"));
    assert_eq!(*vertical_alignment, Some(VerticalAlignment::Bottom));
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(paragraphs[0].runs.len(), 2);
    assert_eq!(
        paragraphs[0].runs[0].font_family.as_deref(),
        Some("Aptos Display")
    );
    assert_eq!(paragraphs[0].runs[0].font_size_pt, Some(27.0));
    assert_eq!(paragraphs[0].runs[0].bold, Some(true));
    assert_eq!(paragraphs[0].runs[0].underline, Some(true));
    assert_eq!(paragraphs[0].runs[0].strike, Some(true));
    assert_eq!(paragraphs[0].runs[0].baseline, Some(30.0));
    assert_eq!(paragraphs[0].runs[0].language.as_deref(), Some("en-US"));
    assert_eq!(
        paragraphs[0].runs[0].hyperlink.as_deref(),
        Some("https://example.com/docs")
    );
    assert_eq!(
        paragraphs[0].runs[0]
            .color
            .as_ref()
            .map(|color| color.value.as_str()),
        Some("#123456")
    );
    assert_eq!(paragraphs[0].alignment.as_deref(), Some("center"));
    assert_eq!(paragraphs[0].level, Some(1));
    assert_eq!(
        paragraphs[0]
            .bullet
            .as_ref()
            .map(|bullet| bullet.kind.as_str()),
        Some("character")
    );
    assert_eq!(
        paragraphs[0].line_spacing,
        Some(TextSpacing {
            value: 18.0,
            unit: TextSpacingUnit::Points,
        })
    );
    assert_eq!(paragraphs[0].rtl, Some(true));
    assert_eq!(paragraphs[0].runs[1].italic, Some(true));

    let SlideNode::Image {
        transform,
        asset_id,
        effects,
        ..
    } = &slide.nodes[1]
    else {
        panic!("expected image")
    };
    assert_eq!(transform.flip_vertical, Some(true));
    assert_eq!(asset_id, "ppt/media/image1.png");
    assert_eq!(document.assets[asset_id].content_type, "image/png");
    assert_eq!(
        effects.as_ref().and_then(|value| value.bi_level_threshold),
        Some(0.6)
    );

    let SlideNode::Group {
        transform,
        child_transform,
        children,
        ..
    } = &slide.nodes[2]
    else {
        panic!("expected group")
    };
    assert_eq!((transform.x, transform.y), (10_000, 11_000));
    assert_eq!(
        child_transform
            .as_ref()
            .map(|value| (value.x, value.y, value.width, value.height)),
        Some((10, 20, 100, 200))
    );
    assert!(matches!(children.as_slice(), [SlideNode::Shape { .. }]));

    let SlideNode::Table {
        rows,
        column_widths,
        row_heights,
        ..
    } = &slide.nodes[3]
    else {
        panic!("expected table")
    };
    assert_eq!(column_widths, &[1000, 2000]);
    assert_eq!(row_heights, &[500, 600]);
    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0].len(), 1, "merged continuation cell is omitted");
    assert_eq!(rows[0][0].col_span, Some(2));
    assert_eq!(rows[0][0].paragraphs[0].runs[0].text, "Header");
    assert!(rows[0][0].borders.contains_key("top"));

    let SlideNode::Chart {
        chart_type,
        title,
        series,
        has_legend,
        chart_xml,
        chart_style_xml,
        chart_colors_xml,
        ..
    } = &slide.nodes[4]
    else {
        panic!("expected chart")
    };
    assert_eq!(chart_type, "bar");
    assert!(
        chart_xml
            .as_deref()
            .is_some_and(|xml| xml.contains("<c:barChart>")),
        "raw chart XML should be preserved for high-fidelity rendering"
    );
    assert!(chart_style_xml
        .as_deref()
        .is_some_and(|xml| xml.contains("chartStyle")));
    assert!(chart_colors_xml
        .as_deref()
        .is_some_and(|xml| xml.contains("colorStyle")));
    assert_eq!(title.as_deref(), Some("Performance"));
    assert_eq!(*has_legend, Some(true));
    assert_eq!(series.len(), 2);
    assert_eq!(series[0].name.as_deref(), Some("Baseline"));
    assert_eq!(series[0].values, vec![Some(700.0), None, Some(280.0)]);
    assert_eq!(
        series[0].categories.as_ref().unwrap(),
        &[
            serde_json::Value::String("001".into()),
            serde_json::Value::String("Render".into())
        ]
    );
    assert_eq!(series[0].color.as_ref().unwrap().value, "#D8663A");
}

#[test]
fn resolves_chart_parts_from_relationship_type_outside_the_standard_chart_folder() {
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
        "ppt/slides/_rels/slide1.xml.rels",
        br#"<Relationships><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/slide1.xml",
        br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:c="c">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr/>
            <p:graphicFrame>
              <p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/></p:nvGraphicFramePr>
              <p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></p:xfrm>
              <a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
            </p:graphicFrame>
          </p:spTree></p:cSld>
        </p:sld>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/charts/chart1.xml",
        br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:ser>
          <c:tx><c:v>Series</c:v></c:tx>
          <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
          <c:val><c:numLit><c:pt idx="0"><c:v>42</c:v></c:pt></c:numLit></c:val>
        </c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
    );
    zip.finish().unwrap();
    let bytes = output.into_inner();

    let document = pptx_core::parse_presentation(&bytes, &ParseLimits::default()).unwrap();
    let SlideNode::Chart {
        chart_type,
        series,
        chart_xml,
        ..
    } = &document.slides[0].nodes[0]
    else {
        panic!("expected chart node")
    };
    assert_eq!(chart_type, "bar");
    assert_eq!(series[0].values, vec![Some(42.0)]);
    assert!(chart_xml
        .as_deref()
        .is_some_and(|xml| xml.contains("barChart")));
}

#[test]
fn resolves_chartex_frames_wrapped_in_alternate_content() {
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
        "ppt/slides/_rels/slide1.xml.rels",
        br#"<Relationships><Relationship Id="rIdChartEx" Target="../charts/chartEx1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/slide1.xml",
        br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:mc="mc" xmlns:cx="cx">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
            <p:grpSpPr/>
            <mc:AlternateContent>
              <mc:Choice Requires="cx1">
                <p:graphicFrame>
                  <p:nvGraphicFramePr><p:cNvPr id="2" name="Funnel"/></p:nvGraphicFramePr>
                  <p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></p:xfrm>
                  <a:graphic><a:graphicData uri="chartex"><cx:chart r:id="rIdChartEx"/></a:graphicData></a:graphic>
                </p:graphicFrame>
              </mc:Choice>
              <mc:Fallback>
                <p:sp><p:nvSpPr><p:cNvPr id="3" name="Fallback"/></p:nvSpPr><p:spPr/></p:sp>
              </mc:Fallback>
            </mc:AlternateContent>
          </p:spTree></p:cSld>
        </p:sld>"#,
    );
    add_file(
        &mut zip,
        "ppt/charts/chartEx1.xml",
        br#"<cx:chartSpace xmlns:cx="cx" xmlns:a="a">
          <cx:chartData><cx:data id="0">
            <cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim>
            <cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">4</cx:pt><cx:pt idx="1">2</cx:pt></cx:lvl></cx:numDim>
          </cx:data></cx:chartData>
          <cx:chart><cx:plotArea><cx:plotAreaRegion>
            <cx:series layoutId="funnel"><cx:dataId val="0"/></cx:series>
          </cx:plotAreaRegion></cx:plotArea></cx:chart>
        </cx:chartSpace>"#,
    );
    zip.finish().unwrap();
    let bytes = output.into_inner();

    let document = pptx_core::parse_presentation(&bytes, &ParseLimits::default()).unwrap();
    let slide = &document.slides[0];
    assert_eq!(slide.nodes.len(), 1, "the Choice branch should win");
    let SlideNode::Chart {
        chart_type,
        chart_xml,
        ..
    } = &slide.nodes[0]
    else {
        panic!("expected chartEx chart node, got {:?}", slide.nodes[0])
    };
    assert_eq!(chart_type, "chartEx");
    assert!(chart_xml
        .as_deref()
        .is_some_and(|xml| xml.contains("funnel")));
}

fn slide_only_pptx(slide_xml: &str) -> Vec<u8> {
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
    add_file(&mut zip, "ppt/slides/slide1.xml", slide_xml.as_bytes());
    zip.finish().unwrap();
    output.into_inner()
}

fn inherited_theme_pptx() -> Vec<u8> {
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
        br#"<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rIdMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/_rels/slide1.xml.rels",
        br#"<Relationships><Relationship Id="rIdLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
        br#"<Relationships><Relationship Id="rIdMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        br#"<Relationships><Relationship Id="rIdTheme" Target="../theme/theme1.xml"/><Relationship Id="rIdLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/slide1.xml",
        br#"<p:sld xmlns:p="p" xmlns:a="a" show="0">
          <p:cSld name="Inherited slide"><p:spTree>
            <p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
              <p:spPr/>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Inherited title</a:t></a:r></a:p></p:txBody>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:ln><a:noFill/></a:ln></p:spPr>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Inherited body</a:t></a:r></a:p><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>No bullet</a:t></a:r></a:p></p:txBody>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="4" name="Unknown scheme"/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm><a:solidFill><a:schemeClr val="notAThemeSlot"/></a:solidFill></p:spPr>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sld>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideLayouts/slideLayout1.xml",
        br#"<p:sldLayout xmlns:p="p" xmlns:a="a">
          <p:cSld name="Title and Content"><p:spTree>
            <p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="Layout title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="6000000" cy="900000"/></a:xfrm></p:spPr>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="3" name="Layout body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="100000" y="1200000"/><a:ext cx="6000000" cy="4000000"/></a:xfrm><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></p:spPr>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sldLayout>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideMasters/slideMaster1.xml",
        br#"<p:sldMaster xmlns:p="p" xmlns:a="a">
          <p:cSld name="Fixture master">
            <p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:bgPr></p:bg>
            <p:spTree>
              <p:nvGrpSpPr/><p:grpSpPr/>
              <p:sp><p:nvSpPr><p:cNvPr id="2" name="Master title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:sp>
              <p:sp><p:nvSpPr><p:cNvPr id="3" name="Master body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="5" y="6"/><a:ext cx="7" cy="8"/></a:xfrm><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></p:spPr></p:sp>
            </p:spTree>
          </p:cSld>
          <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"/>
          <p:txStyles>
            <p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="3200"><a:latin typeface="+mj-lt"/><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle>
            <p:bodyStyle><a:lvl1pPr><a:buChar char="*"/><a:defRPr sz="1800"><a:latin typeface="+mn-lt"/><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:bodyStyle>
            <p:otherStyle><a:lvl1pPr><a:defRPr/></a:lvl1pPr></p:otherStyle>
          </p:txStyles>
        </p:sldMaster>"#,
    );
    add_file(
        &mut zip,
        "ppt/theme/theme1.xml",
        br#"<a:theme xmlns:a="a" name="Fixture theme"><a:themeElements>
          <a:clrScheme name="Fixture colors">
            <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
            <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
            <a:accent1><a:srgbClr val="336699"/></a:accent1><a:accent2><a:srgbClr val="AABBCC"/></a:accent2>
          </a:clrScheme>
          <a:fontScheme name="Fixture fonts"><a:majorFont><a:latin typeface="Theme Heading"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Theme Body"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
        </a:themeElements></a:theme>"#,
    );
    zip.finish().unwrap();
    output.into_inner()
}

#[test]
fn resolves_layout_master_theme_and_placeholder_inheritance() {
    let document =
        pptx_core::parse_presentation(&inherited_theme_pptx(), &ParseLimits::default()).unwrap();
    assert_eq!(document.layouts.len(), 1);
    assert_eq!(document.masters.len(), 1);
    assert_eq!(document.themes.len(), 1);
    assert_eq!(document.themes[0].colors["accent1"], "#336699");
    assert_eq!(document.themes[0].major_fonts["latin"], "Theme Heading");
    assert_eq!(document.themes[0].minor_fonts["latin"], "Theme Body");

    let slide = &document.slides[0];
    assert_eq!(slide.name.as_deref(), Some("Inherited slide"));
    assert_eq!(slide.hidden, Some(true));
    assert_eq!(
        slide.layout_id.as_deref(),
        Some("ppt/slideLayouts/slideLayout1.xml")
    );
    assert_eq!(
        slide.master_id.as_deref(),
        Some("ppt/slideMasters/slideMaster1.xml")
    );
    assert!(
        matches!(&slide.background, Some(FillStyle::Solid { color }) if color.value == "#FFFFFF")
    );

    let SlideNode::Shape {
        transform,
        paragraphs,
        ..
    } = &slide.nodes[0]
    else {
        panic!("expected title placeholder")
    };
    assert_eq!((transform.x, transform.y), (100_000, 200_000));
    assert_eq!((transform.width, transform.height), (6_000_000, 900_000));
    assert_eq!(paragraphs[0].alignment.as_deref(), Some("center"));
    assert_eq!(
        paragraphs[0].runs[0].font_family.as_deref(),
        Some("Theme Heading")
    );
    assert_eq!(paragraphs[0].runs[0].font_size_pt, Some(32.0));
    assert_eq!(
        paragraphs[0].runs[0].color.as_ref().unwrap().value,
        "#E8EAEF"
    );

    let SlideNode::Shape {
        transform,
        fill,
        line,
        paragraphs,
        ..
    } = &slide.nodes[1]
    else {
        panic!("expected body placeholder")
    };
    assert_eq!((transform.x, transform.y), (100_000, 1_200_000));
    assert_eq!((transform.width, transform.height), (6_000_000, 4_000_000));
    assert!(matches!(fill, Some(FillStyle::Solid { color }) if color.value == "#AABBCC"));
    assert!(
        line.is_none(),
        "an explicit noFill line clears the master line"
    );
    assert_eq!(
        paragraphs[0].runs[0].font_family.as_deref(),
        Some("Theme Body")
    );
    assert_eq!(
        paragraphs[0].runs[0].color.as_ref().unwrap().value,
        "#000000"
    );
    assert_eq!(
        paragraphs[0]
            .bullet
            .as_ref()
            .map(|bullet| bullet.kind.as_str()),
        Some("character")
    );
    assert!(paragraphs[1].bullet.is_none());

    assert!(matches!(
        &slide.nodes[2],
        SlideNode::Shape { fill: None, .. }
    ));
    assert!(slide
        .warnings
        .iter()
        .any(|warning| warning.feature.as_deref() == Some("theme-color")));
}

fn theme_image_relationship_pptx() -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut output);
    add_file(
        &mut zip,
        "ppt/presentation.xml",
        br#"<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#,
    );
    add_file(
        &mut zip,
        "ppt/_rels/presentation.xml.rels",
        br#"<Relationships><Relationship Id="rIdSlide" Target="slides/slide1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/_rels/slide1.xml.rels",
        br#"<Relationships><Relationship Id="rIdLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
        br#"<Relationships><Relationship Id="rIdMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        br#"<Relationships><Relationship Id="rIdTheme" Target="../theme/theme1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/theme/_rels/theme1.xml.rels",
        br#"<Relationships><Relationship Id="rIdThemeImage" Target="../media/theme-image.png"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/slide1.xml",
        br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
          <p:sp><p:nvSpPr><p:cNvPr id="2" name="Theme image fill"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>
        </p:spTree></p:cSld></p:sld>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideLayouts/slideLayout1.xml",
        br#"<p:sldLayout xmlns:p="p"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideMasters/slideMaster1.xml",
        br#"<p:sldMaster xmlns:p="p" xmlns:a="a"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:clrMap accent1="accent1"/></p:sldMaster>"#,
    );
    add_file(
        &mut zip,
        "ppt/theme/theme1.xml",
        br#"<a:theme xmlns:a="a" xmlns:r="r"><a:themeElements>
          <a:clrScheme name="Theme image colors"><a:accent1><a:srgbClr val="336699"/></a:accent1></a:clrScheme>
          <a:fontScheme name="Theme image fonts"><a:majorFont/><a:minorFont/></a:fontScheme>
          <a:fmtScheme name="Theme image formats">
            <a:fillStyleLst><a:blipFill><a:blip r:embed="rIdThemeImage"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></a:fillStyleLst>
            <a:lnStyleLst/>
            <a:bgFillStyleLst><a:blipFill><a:blip r:embed="rIdThemeImage"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></a:bgFillStyleLst>
          </a:fmtScheme>
        </a:themeElements></a:theme>"#,
    );
    add_file(
        &mut zip,
        "ppt/media/theme-image.png",
        b"\x89PNG\r\n\x1a\ntheme-image",
    );
    zip.finish().unwrap();
    output.into_inner()
}

#[test]
fn resolves_theme_image_relationships_for_style_matrix_fills() {
    let document =
        pptx_core::parse_presentation(&theme_image_relationship_pptx(), &ParseLimits::default())
            .unwrap();
    let slide = &document.slides[0];
    assert!(matches!(
        &slide.background,
        Some(FillStyle::Image {
            asset_id,
            mode: FillImageMode::Stretch,
            ..
        }) if asset_id == "ppt/media/theme-image.png"
    ));
    assert!(matches!(
        &slide.nodes[0],
        SlideNode::Shape {
            fill: Some(FillStyle::Image {
                asset_id,
                mode: FillImageMode::Stretch,
                ..
            }),
            ..
        } if asset_id == "ppt/media/theme-image.png"
    ));
    assert!(document.assets.contains_key("ppt/media/theme-image.png"));
}

#[test]
fn rejects_shallow_wide_xml_over_the_node_budget() {
    let shapes = (0..64)
        .map(|index| {
            format!(
                "<p:sp><p:nvSpPr><p:cNvPr id=\"{}\" name=\"wide\"/></p:nvSpPr><p:spPr/></p:sp>",
                index + 2
            )
        })
        .collect::<String>();
    let slide = format!(
        "<p:sld xmlns:p=\"p\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"
    );
    let error = pptx_core::parse_presentation(
        &slide_only_pptx(&slide),
        &ParseLimits {
            max_xml_nodes: 24,
            ..Default::default()
        },
    )
    .unwrap_err();
    assert!(matches!(error, ParseError::ResourceLimit(message) if message.contains("node count")));
}

#[test]
fn never_exposes_untrusted_preset_colors_as_renderable_values() {
    let slide = r#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="Unsafe color"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:solidFill><a:prstClr val="url(javascript:alert(1))"/></a:solidFill></p:spPr></p:sp>
    </p:spTree></p:cSld></p:sld>"#;
    let document =
        pptx_core::parse_presentation(&slide_only_pptx(slide), &ParseLimits::default()).unwrap();
    assert!(matches!(
        &document.slides[0].nodes[0],
        SlideNode::Shape { fill: None, .. }
    ));
    assert!(document.slides[0]
        .warnings
        .iter()
        .any(|warning| warning.feature.as_deref() == Some("preset-color")));
    assert!(!serde_json::to_string(&document)
        .unwrap()
        .contains("javascript"));
}

fn visual_fidelity_pptx() -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut output);
    add_file(
        &mut zip,
        "ppt/presentation.xml",
        br#"<p:presentation xmlns:p="p" xmlns:a="a" xmlns:r="r">
          <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
          <p:sldSz cx="12192000" cy="6858000"/>
          <p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1100"/></a:lvl1pPr></p:defaultTextStyle>
        </p:presentation>"#,
    );
    add_file(
        &mut zip,
        "ppt/_rels/presentation.xml.rels",
        br#"<Relationships><Relationship Id="rIdSlide" Target="slides/slide1.xml"/><Relationship Id="rIdMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rIdTableStyles" Target="tableStyles.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/_rels/slide1.xml.rels",
        br#"<Relationships>
          <Relationship Id="rIdLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdBackground" Target="../media/background.png"/>
          <Relationship Id="rIdPicture" Target="../media/picture.png"/>
          <Relationship Id="rIdOleFallback" Target="../media/ole.png"/>
          <Relationship Id="rIdShapeFill" Target="../media/shape-fill.png"/>
        </Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
        br#"<Relationships><Relationship Id="rIdMaster" Target="../slideMasters/slideMaster1.xml"/><Relationship Id="rIdLayoutLogo" Target="../media/layout-logo.png"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        br#"<Relationships><Relationship Id="rIdTheme" Target="../theme/theme1.xml"/><Relationship Id="rIdLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rIdMasterLogo" Target="../media/master-logo.png"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/slide1.xml",
        br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r">
          <p:cSld><p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBackground"/><a:tile/></a:blipFill></p:bgPr></p:bg>
          <p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="Styled body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
              <p:spPr/>
              <p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="2"><a:schemeClr val="accent2"/></a:fillRef><a:fontRef idx="major"><a:schemeClr val="tx1"/></a:fontRef></p:style>
              <p:txBody><a:bodyPr lIns="100" wrap="none" numCol="2" spcCol="700" rtlCol="1" spcFirstLastPara="1" horzOverflow="overflow" vertOverflow="ellipsis" rot="5400000"><a:normAutofit fontScale="92000" lnSpcReduction="20000"/></a:bodyPr><a:lstStyle/>
                <a:p><a:pPr marL="500" indent="-100"><a:buFont typeface="Wingdings"/><a:buSzPct val="80000"/><a:buAutoNum type="arabicPeriod" startAt="3"/><a:lnSpc><a:spcPts val="2000"/></a:lnSpc><a:spcBef><a:spcPct val="150000"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r><a:rPr spc="-8"/><a:t>Exact spacing</a:t></a:r></a:p>
              </p:txBody>
            </p:sp>
            <p:pic><p:nvPicPr><p:cNvPr id="3" name="Cropped picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdPicture"><a:alphaModFix amt="65000"/></a:blip><a:srcRect l="10000" t="20000" r="5000" b="2500"/><a:stretch/></p:blipFill><p:spPr><a:xfrm><a:off x="200" y="300"/><a:ext cx="400" cy="500"/></a:xfrm></p:spPr></p:pic>
            <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="OLE fallback"/></p:nvGraphicFramePr><p:xfrm><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></p:xfrm><a:graphic><a:graphicData><mc:AlternateContent xmlns:mc="mc"><mc:Fallback><p:oleObj><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdOleFallback"/><a:stretch/></p:blipFill><p:spPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></a:xfrm></p:spPr></p:pic></p:oleObj></mc:Fallback></mc:AlternateContent></a:graphicData></a:graphic></p:graphicFrame>
            <p:sp><p:nvSpPr><p:cNvPr id="5" name="Image fill"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="5000" y="6000"/><a:ext cx="7000" cy="8000"/></a:xfrm><a:prstGeom prst="rect"/><a:blipFill><a:blip r:embed="rIdShapeFill"/><a:srcRect l="12500"/><a:tile/></a:blipFill></p:spPr></p:sp>
            <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Rotated table"/></p:nvGraphicFramePr><p:xfrm><a:off x="9000" y="10000"/><a:ext cx="11000" cy="12000"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>{VISUAL-STYLE}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="11000"/></a:tblGrid><a:tr h="12000"><a:tc><a:txBody><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody><a:tcPr marL="100" marR="200" marT="300" marB="400" anchor="ctr" vert="vert270"/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
          </p:spTree></p:cSld>
        </p:sld>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideLayouts/slideLayout1.xml",
        br#"<p:sldLayout xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
          <p:pic><p:nvPicPr><p:cNvPr id="8" name="Layout logo"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdLayoutLogo"/><a:stretch/></p:blipFill><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm></p:spPr></p:pic>
          <p:sp><p:nvSpPr><p:cNvPr id="2" name="Layout body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="6000000" cy="4000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr tIns="111"/><a:lstStyle/></p:txBody></p:sp>
        </p:spTree></p:cSld></p:sldLayout>"#,
    );
    add_file(
        &mut zip,
        "ppt/slideMasters/slideMaster1.xml",
        br#"<p:sldMaster xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
          <p:pic><p:nvPicPr><p:cNvPr id="9" name="Master logo"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdMasterLogo"/><a:stretch/></p:blipFill><p:spPr><a:xfrm><a:off x="50" y="60"/><a:ext cx="70" cy="80"/></a:xfrm></p:spPr></p:pic>
          <p:sp><p:nvSpPr><p:cNvPr id="2" name="Master body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr rIns="222" anchor="b"/><a:lstStyle/></p:txBody></p:sp>
        </p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" accent1="accent1" accent2="accent2"/><p:txStyles><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1400"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>"#,
    );
    add_file(
        &mut zip,
        "ppt/theme/theme1.xml",
        br#"<a:theme xmlns:a="a" name="Visual theme"><a:themeElements>
          <a:clrScheme name="Visual"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="336699"/></a:accent1><a:accent2><a:srgbClr val="CC6600"/></a:accent2></a:clrScheme>
          <a:fontScheme name="Visual fonts"><a:majorFont><a:latin typeface="Theme Heading"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Theme Body"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
          <a:fmtScheme name="Visual formats"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln></a:lnStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
        </a:themeElements></a:theme>"#,
    );
    add_file(
        &mut zip,
        "ppt/tableStyles.xml",
        br#"<a:tblStyleLst xmlns:a="a" def="{VISUAL-STYLE}"><a:tblStyle styleId="{VISUAL-STYLE}" styleName="Visual"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:ln w="9525"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></a:left></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent2"><a:tint val="50000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a:firstRow><a:tcTxStyle b="1"><a:fontRef idx="major"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:bottom><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle></a:tblStyleLst>"#,
    );
    for path in [
        "ppt/media/background.png",
        "ppt/media/picture.png",
        "ppt/media/ole.png",
        "ppt/media/shape-fill.png",
        "ppt/media/layout-logo.png",
        "ppt/media/master-logo.png",
    ] {
        add_file(&mut zip, path, b"\x89PNG\r\n\x1a\nvisual");
    }
    zip.finish().unwrap();
    output.into_inner()
}

#[test]
fn resolves_visual_fidelity_inheritance_images_and_text_metrics() {
    let document =
        pptx_core::parse_presentation(&visual_fidelity_pptx(), &ParseLimits::default()).unwrap();
    let slide = &document.slides[0];
    assert!(matches!(
        &slide.background,
        Some(FillStyle::Image { asset_id, mode: FillImageMode::Tile, .. })
            if asset_id == "ppt/media/background.png"
    ));
    assert_eq!(slide.nodes.len(), 7);
    assert!(matches!(
        &slide.nodes[0],
        SlideNode::Image { id, asset_id, .. }
            if id.starts_with("master:") && asset_id == "ppt/media/master-logo.png"
    ));
    assert!(matches!(
        &slide.nodes[1],
        SlideNode::Image { id, asset_id, .. }
            if id.starts_with("layout:") && asset_id == "ppt/media/layout-logo.png"
    ));

    let SlideNode::Shape {
        transform,
        fill,
        line,
        paragraphs,
        vertical_alignment,
        text_insets,
        autofit,
        text_rotation,
        text_wrap,
        column_count,
        column_spacing,
        right_to_left_columns,
        space_first_last_paragraph,
        ..
    } = &slide.nodes[2]
    else {
        panic!("expected styled body shape")
    };
    assert_eq!((transform.x, transform.y), (100_000, 200_000));
    assert!(matches!(fill, Some(FillStyle::Gradient { stops, .. }) if stops.len() == 2));
    assert_eq!(line.as_ref().and_then(|line| line.width), Some(2.0));
    assert_eq!(*vertical_alignment, Some(VerticalAlignment::Bottom));
    assert_eq!(text_insets.unwrap().left, 100);
    assert_eq!(text_insets.unwrap().top, 111);
    assert_eq!(text_insets.unwrap().right, 222);
    assert_eq!(text_insets.unwrap().bottom, 45_720);
    assert_eq!(autofit.unwrap().mode, TextAutoFitMode::Normal);
    assert_eq!(autofit.unwrap().font_scale, Some(0.92));
    assert_eq!(autofit.unwrap().line_spacing_reduction, Some(0.2));
    assert_eq!(*text_rotation, Some(90.0));
    assert_eq!(text_wrap.as_deref(), Some("none"));
    assert_eq!(*column_count, Some(2));
    assert_eq!(*column_spacing, Some(700));
    assert_eq!(*right_to_left_columns, Some(true));
    assert_eq!(*space_first_last_paragraph, Some(true));
    assert_eq!(paragraphs[0].margin_left_emu, Some(500));
    assert_eq!(paragraphs[0].indent_emu, Some(-100));
    assert_eq!(
        paragraphs[0].line_spacing.unwrap().unit,
        TextSpacingUnit::Points
    );
    assert_eq!(paragraphs[0].line_spacing.unwrap().value, 20.0);
    assert_eq!(
        paragraphs[0].space_before.unwrap().unit,
        TextSpacingUnit::Percent
    );
    assert_eq!(paragraphs[0].space_before.unwrap().value, 1.5);
    assert_eq!(paragraphs[0].space_after.unwrap().value, 6.0);
    let bullet = paragraphs[0].bullet.as_ref().unwrap();
    assert_eq!(bullet.font_family.as_deref(), Some("Wingdings"));
    assert_eq!(bullet.size_percent, Some(0.8));
    assert_eq!(bullet.start_at, Some(3));
    assert_eq!(
        paragraphs[0].runs[0].font_family.as_deref(),
        Some("Theme Heading")
    );
    assert_eq!(paragraphs[0].runs[0].font_size_pt, Some(14.0));
    assert_eq!(paragraphs[0].runs[0].character_spacing_pt, Some(-0.08));

    assert!(matches!(
        &slide.nodes[3],
        SlideNode::Image {
            crop: Some(crop),
            opacity: Some(0.65),
            preserve_aspect_ratio: false,
            ..
        } if crop.left == 0.1 && crop.top == 0.2 && crop.right == 0.05 && crop.bottom == 0.025
    ));
    assert!(matches!(
        &slide.nodes[4],
        SlideNode::Image { name, asset_id, .. }
            if name == "OLE fallback" && asset_id == "ppt/media/ole.png"
    ));
    assert!(matches!(
        &slide.nodes[5],
        SlideNode::Shape {
            fill: Some(FillStyle::Image {
                asset_id,
                mode: FillImageMode::Tile,
                crop: Some(crop),
                ..
            }),
            ..
        } if asset_id == "ppt/media/shape-fill.png" && crop.left == 0.125
    ));
    let SlideNode::Table { rows, .. } = &slide.nodes[6] else {
        panic!("expected table")
    };
    assert_eq!(rows[0][0].text_insets.unwrap().left, 100);
    assert_eq!(rows[0][0].text_insets.unwrap().bottom, 400);
    assert_eq!(
        rows[0][0].vertical_alignment,
        Some(VerticalAlignment::Middle)
    );
    assert_eq!(rows[0][0].text_rotation, Some(270.0));
    assert!(matches!(
        &rows[0][0].fill,
        Some(FillStyle::Solid { color }) if color.value == "#336699"
    ));
    assert_eq!(rows[0][0].paragraphs[0].runs[0].bold, Some(true));
    assert_eq!(
        rows[0][0].paragraphs[0].runs[0].font_family.as_deref(),
        Some("Theme Heading")
    );
    assert_eq!(
        rows[0][0].paragraphs[0].runs[0]
            .color
            .as_ref()
            .map(|color| color.value.as_str()),
        Some("#FFFFFF")
    );
    assert!(rows[0][0].borders.contains_key("left"));
    assert_eq!(
        rows[0][0].borders.get("bottom").and_then(|line| line.width),
        Some(2.0)
    );
}

fn smartart_pptx(data_xml: &[u8], drawing_xml: Option<&[u8]>) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut output);
    add_file(
        &mut zip,
        "ppt/presentation.xml",
        br#"<p:presentation xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#,
    );
    add_file(
        &mut zip,
        "ppt/_rels/presentation.xml.rels",
        br#"<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>"#,
    );
    add_file(
        &mut zip,
        "ppt/slides/slide1.xml",
        br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:dgm="dgm"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Diagram 1"/></p:nvGraphicFramePr><p:xfrm><a:off x="100" y="200"/><a:ext cx="1000000" cy="600000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#,
    );
    let drawing_relationship = if drawing_xml.is_some() {
        r#"<Relationship Id="rId6" Target="../diagrams/drawing1.xml"/>"#
    } else {
        ""
    };
    add_file(
        &mut zip,
        "ppt/slides/_rels/slide1.xml.rels",
        format!(
            r#"<Relationships><Relationship Id="rId2" Target="../diagrams/data1.xml"/>{drawing_relationship}</Relationships>"#
        )
        .as_bytes(),
    );
    add_file(&mut zip, "ppt/diagrams/data1.xml", data_xml);
    if let Some(drawing_xml) = drawing_xml {
        add_file(&mut zip, "ppt/diagrams/drawing1.xml", drawing_xml);
    }
    zip.finish().unwrap();
    output.into_inner()
}

#[test]
fn materialized_smartart_uses_drawing_geometry_and_semantic_text_mapping() {
    let data = br#"<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a" xmlns:dsp="dsp"><dgm:ptLst>
      <dgm:pt modelId="semantic-1"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1600"/><a:t>Mapped label</a:t></a:r></a:p></dgm:t></dgm:pt>
      <dgm:pt modelId="presentation-1" type="pres"><dgm:prSet presAssocID="semantic-1"/><dgm:spPr/></dgm:pt>
    </dgm:ptLst><dgm:extLst><a:ext uri="diagram"><dsp:dataModelExt relId="rId6"/></a:ext></dgm:extLst></dgm:dataModel>"#;
    let drawing = br#"<dsp:drawing xmlns:dsp="dsp" xmlns:a="a"><dsp:spTree><dsp:nvGrpSpPr/><dsp:grpSpPr/>
      <dsp:sp modelId="presentation-1"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm rot="60000"><a:off x="10000" y="20000"/><a:ext cx="400000" cy="200000"/></a:xfrm><a:prstGeom prst="roundRect"/><a:solidFill><a:srgbClr val="336699"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr sz="1600"/></a:p></dsp:txBody></dsp:sp>
    </dsp:spTree></dsp:drawing>"#;
    let document =
        pptx_core::parse_presentation(&smartart_pptx(data, Some(drawing)), &ParseLimits::default())
            .unwrap();
    let SlideNode::Group {
        transform,
        children,
        child_transform,
        ..
    } = &document.slides[0].nodes[0]
    else {
        panic!("expected materialized SmartArt group")
    };
    assert_eq!((transform.x, transform.y), (100, 200));
    assert_eq!((transform.width, transform.height), (1_000_000, 600_000));
    assert_eq!(child_transform.as_ref().unwrap().width, 1_000_000);
    assert_eq!(children.len(), 1);
    let SlideNode::Shape {
        id,
        transform,
        geometry,
        fill,
        paragraphs,
        ..
    } = &children[0]
    else {
        panic!("expected materialized SmartArt shape")
    };
    assert!(id.contains("presentation-1"));
    assert_eq!((transform.x, transform.y), (10_000, 20_000));
    assert_eq!(transform.rotation, Some(1.0));
    assert_eq!(geometry.preset.as_deref(), Some("roundRect"));
    assert!(matches!(fill, Some(FillStyle::Solid { color }) if color.value == "#336699"));
    assert_eq!(paragraphs[0].runs[0].text, "Mapped label");
    assert!(!document.slides[0]
        .warnings
        .iter()
        .any(|warning| warning.feature.as_deref() == Some("smartart")));
}

#[test]
fn smartart_without_drawing_emits_readable_semantic_fallback() {
    let data = br#"<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a"><dgm:ptLst>
      <dgm:pt modelId="doc" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Do not render</a:t></a:r></a:p></dgm:t></dgm:pt>
      <dgm:pt modelId="alpha"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></dgm:t></dgm:pt>
      <dgm:pt modelId="pres" type="pres"><dgm:prSet presAssocID="alpha"/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Presentation helper</a:t></a:r></a:p></dgm:t></dgm:pt>
      <dgm:pt modelId="par" type="parTrans"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Parent helper</a:t></a:r></a:p></dgm:t></dgm:pt>
      <dgm:pt modelId="sib" type="sibTrans"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Sibling helper</a:t></a:r></a:p></dgm:t></dgm:pt>
      <dgm:pt modelId="beta"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Beta</a:t></a:r></a:p></dgm:t></dgm:pt>
    </dgm:ptLst></dgm:dataModel>"#;
    let document =
        pptx_core::parse_presentation(&smartart_pptx(data, None), &ParseLimits::default()).unwrap();
    let SlideNode::Group {
        children,
        child_transform,
        ..
    } = &document.slides[0].nodes[0]
    else {
        panic!("expected semantic SmartArt fallback group")
    };
    assert_eq!(children.len(), 2);
    assert_eq!(child_transform.as_ref().unwrap().height, 600_000);
    let labels = children
        .iter()
        .filter_map(|node| match node {
            SlideNode::Shape { paragraphs, .. } => Some(
                paragraphs
                    .iter()
                    .flat_map(|paragraph| paragraph.runs.iter())
                    .map(|run| run.text.as_str())
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(labels, ["Alpha", "Beta"]);
    assert!(document.slides[0].warnings.iter().any(|warning| {
        warning.code == "degraded-rendering" && warning.feature.as_deref() == Some("smartart")
    }));
}
