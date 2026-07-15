use std::io::{Cursor, Seek, Write};

use pptx_core::{FillStyle, ParseError, ParseLimits, SlideNode, VerticalAlignment};
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
                <p:blipFill><a:blip r:embed="rIdImage"/><a:stretch/></p:blipFill>
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
    assert_eq!(paragraphs[0].line_spacing, Some(18.0));
    assert_eq!(paragraphs[0].rtl, Some(true));
    assert_eq!(paragraphs[0].runs[1].italic, Some(true));

    let SlideNode::Image {
        transform,
        asset_id,
        ..
    } = &slide.nodes[1]
    else {
        panic!("expected image")
    };
    assert_eq!(transform.flip_vertical, Some(true));
    assert_eq!(asset_id, "ppt/media/image1.png");
    assert_eq!(document.assets[asset_id].content_type, "image/png");

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
        ..
    } = &slide.nodes[4]
    else {
        panic!("expected chart")
    };
    assert_eq!(chart_type, "bar");
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
        "#5C85AD"
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
