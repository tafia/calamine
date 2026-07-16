// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Parser for DrawingML chart parts (`xl/charts/chartN.xml`) and for the
//! chart references inside spreadsheet drawings (`xl/drawings/drawingN.xml`).

use std::io::BufRead;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;

use crate::attrs::{decode_attr, RawAttributes};
use crate::chart::{
    Chart, ChartAxis, ChartAxisCrosses, ChartAxisPosition, ChartAxisType, ChartBar3dShape,
    ChartCellAnchor, ChartCrossBetween, ChartDataLabel, ChartDataLabelPosition, ChartDataLabels,
    ChartDataPoint, ChartDataSource, ChartDataTable, ChartDisplayBlanksAs, ChartDisplayUnits,
    ChartEditAs, ChartErrorBars, ChartErrorBarsDirection, ChartErrorBarsType,
    ChartErrorBarsValueType, ChartExLayout, ChartExParentLabelLayout, ChartExQuartileMethod,
    ChartFill, ChartFormat, ChartGradientStop, ChartGroup, ChartLegend, ChartLegendPosition,
    ChartLine, ChartLineDashType, ChartLines, ChartMarker, ChartMarkerType, ChartOfPieSplitType,
    ChartPosition, ChartSeries, ChartSizeRepresents, ChartTickLabelPosition, ChartTickMark,
    ChartTitle, ChartTrendline, ChartTrendlineType, ChartType, ChartUpDownBars, ChartView3d,
};
use crate::datatype::Data;
use crate::style::{Color, Font, FontStyle, FontWeight, RichText, TextRun, UnderlineStyle};
use crate::utils::unescape_entity_to_buffer;

use super::XlsxError;

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

/// Consume the rest of the subtree of the element `e` (whose `Start` event
/// was just read).
///
/// Note: the readers created by `xml_reader` expand empty elements, so every
/// element is guaranteed to produce a matching `End` event.
fn skip_element<RS: BufRead>(xml: &mut XmlReader<RS>, e: &BytesStart) -> Result<(), XlsxError> {
    let name = e.name().as_ref().to_vec();
    let mut depth = 0u32;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(s) if s.name().as_ref() == name.as_slice() => depth += 1,
            Event::End(end) if end.name().as_ref() == name.as_slice() => {
                if depth == 0 {
                    return Ok(());
                }
                depth -= 1;
            }
            Event::Eof => return Err(XlsxError::XmlEof("chart element")),
            _ => (),
        }
    }
}

/// Read the text content of the element `e` up to its `End` event.
fn read_text<RS: BufRead>(xml: &mut XmlReader<RS>, e: &BytesStart) -> Result<String, XlsxError> {
    let mut buf = Vec::new();
    let mut value = String::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Text(t) => value.push_str(&t.xml10_content()?),
            Event::CData(t) => value.push_str(&t.xml10_content()?),
            Event::GeneralRef(r) => unescape_entity_to_buffer(&r, &mut value)?,
            Event::End(end) if end.name() == e.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("chart text")),
            _ => (),
        }
    }
    Ok(value)
}

fn parse_bytes<T: std::str::FromStr>(bytes: &[u8]) -> Option<T> {
    std::str::from_utf8(bytes).ok()?.trim().parse().ok()
}

/// Read the `val` attribute of `e` parsed as `T`.
fn val_attr<T: std::str::FromStr>(e: &BytesStart) -> Result<Option<T>, XlsxError> {
    Ok(e.raw_attr(b"val")?.and_then(parse_bytes))
}

/// Read the `val` attribute of `e` as an owned string.
fn val_attr_string<RS: BufRead>(
    xml: &XmlReader<RS>,
    e: &BytesStart,
) -> Result<Option<String>, XlsxError> {
    match e.raw_attr(b"val")? {
        Some(v) => Ok(Some(decode_attr(&xml.decoder(), v)?)),
        None => Ok(None),
    }
}

/// Read the `val` attribute of `e` as a boolean (`1`/`true`).
fn val_attr_bool(e: &BytesStart) -> Result<Option<bool>, XlsxError> {
    Ok(e.raw_attr(b"val")?
        .map(|v| matches!(v, b"1" | b"true" | b"on")))
}

// ---------------------------------------------------------------------------
// Chart space
// ---------------------------------------------------------------------------

/// Parse a `c:chartSpace` document into a [`Chart`].
pub(crate) fn parse_chart_space<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    theme_colors: &[Color],
) -> Result<Chart, XlsxError> {
    let mut chart = Chart::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"chartSpace" => (),
                b"chart" => parse_chart(xml, &e, &mut chart, theme_colors)?,
                b"style" => {
                    if chart.style.is_none() {
                        // `c14:style` (inside mc:AlternateContent) stores the
                        // style number offset by 100 relative to `c:style`.
                        chart.style = val_attr::<u32>(&e)?.map(|v| if v > 100 { v - 100 } else { v });
                    }
                    skip_element(xml, &e)?;
                }
                b"spPr" => chart.format = parse_shape_properties(xml, &e, theme_colors)?,
                b"roundedCorners" => {
                    chart.rounded_corners = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                b"date1904" => {
                    chart.date_1904 = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                // Descend into mc:AlternateContent to reach the style value.
                b"AlternateContent" | b"Choice" | b"Fallback" => (),
                _ => skip_element(xml, &e)?,
            },
            Event::Eof => break,
            _ => (),
        }
    }
    Ok(chart)
}

/// Parse the `c:chart` element.
fn parse_chart<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    chart: &mut Chart,
    theme_colors: &[Color],
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    let mut auto_title_deleted = false;
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"title" => chart.title = Some(parse_title(xml, &e, theme_colors)?),
                b"autoTitleDeleted" => {
                    auto_title_deleted = val_attr_bool(&e)?.unwrap_or(false);
                    chart.auto_title_deleted = Some(auto_title_deleted);
                    skip_element(xml, &e)?;
                }
                b"view3D" => chart.view_3d = Some(parse_view_3d(xml, &e)?),
                b"plotArea" => parse_plot_area(xml, &e, chart, theme_colors)?,
                b"legend" => chart.legend = Some(parse_legend(xml, &e, theme_colors)?),
                b"plotVisOnly" => {
                    chart.plot_visible_only = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                b"showDLblsOverMax" => {
                    chart.show_data_labels_over_max = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                b"dispBlanksAs" => {
                    chart.display_blanks_as = match e.raw_attr(b"val")? {
                        Some(b"gap") => Some(ChartDisplayBlanksAs::Gap),
                        Some(b"span") => Some(ChartDisplayBlanksAs::Span),
                        Some(b"zero") => Some(ChartDisplayBlanksAs::Zero),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("chart")),
            _ => (),
        }
    }
    if auto_title_deleted {
        chart.title = None;
    }
    Ok(())
}

/// Parse the `c:plotArea` element: plot groups, axes and formatting.
fn parse_plot_area<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    chart: &mut Chart,
    theme_colors: &[Color],
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"barChart" | b"bar3DChart" | b"lineChart" | b"line3DChart" | b"pieChart"
                | b"pie3DChart" | b"doughnutChart" | b"areaChart" | b"area3DChart"
                | b"scatterChart" | b"radarChart" | b"stockChart" | b"surfaceChart"
                | b"surface3DChart" | b"bubbleChart" | b"ofPieChart" => {
                    let group = parse_chart_group(xml, &e, theme_colors)?;
                    chart.groups.push(group);
                }
                b"catAx" => chart
                    .axes
                    .push(parse_axis(xml, &e, ChartAxisType::Category, theme_colors)?),
                b"valAx" => chart
                    .axes
                    .push(parse_axis(xml, &e, ChartAxisType::Value, theme_colors)?),
                b"dateAx" => chart
                    .axes
                    .push(parse_axis(xml, &e, ChartAxisType::Date, theme_colors)?),
                b"serAx" => chart
                    .axes
                    .push(parse_axis(xml, &e, ChartAxisType::Series, theme_colors)?),
                b"spPr" => chart.plot_area_format = parse_shape_properties(xml, &e, theme_colors)?,
                b"dTable" => chart.data_table = Some(parse_data_table(xml, &e, theme_colors)?),
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("plotArea")),
            _ => (),
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Plot groups and chart type resolution
// ---------------------------------------------------------------------------

/// Parse one plot group element such as `c:barChart` or `c:surface3DChart`.
fn parse_chart_group<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartGroup, XlsxError> {
    let group_tag = parent.local_name().as_ref().to_vec();
    let mut group = ChartGroup::default();

    let mut bar_dir: Option<String> = None;
    let mut grouping: Option<String> = None;
    let mut scatter_style: Option<String> = None;
    let mut radar_style: Option<String> = None;
    let mut of_pie_type: Option<String> = None;
    let mut wireframe = false;

    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"ser" => {
                    let series = parse_series(xml, &e, theme_colors)?;
                    group.series.push(series);
                }
                b"barDir" => {
                    bar_dir = val_attr_string(xml, &e)?;
                    skip_element(xml, &e)?;
                }
                b"grouping" => {
                    grouping = val_attr_string(xml, &e)?;
                    skip_element(xml, &e)?;
                }
                b"scatterStyle" => {
                    scatter_style = val_attr_string(xml, &e)?;
                    skip_element(xml, &e)?;
                }
                b"radarStyle" => {
                    radar_style = val_attr_string(xml, &e)?;
                    skip_element(xml, &e)?;
                }
                b"ofPieType" => {
                    of_pie_type = val_attr_string(xml, &e)?;
                    skip_element(xml, &e)?;
                }
                b"wireframe" => {
                    wireframe = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"gapWidth" => {
                    group.gap_width = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"gapDepth" => {
                    group.gap_depth = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"overlap" => {
                    group.overlap = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"holeSize" => {
                    group.hole_size = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"firstSliceAng" => {
                    group.first_slice_angle = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"bubbleScale" => {
                    group.bubble_scale = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"varyColors" => {
                    group.vary_colors = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                b"axId" => {
                    if let Some(id) = val_attr::<u32>(&e)? {
                        group.axis_ids.push(id);
                    }
                    skip_element(xml, &e)?;
                }
                b"marker" => {
                    group.show_marker = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                b"dLbls" => group.data_labels = Some(parse_data_labels(xml, &e, theme_colors)?),
                b"shape" => {
                    group.shape = match e.raw_attr(b"val")? {
                        Some(b"box") => Some(ChartBar3dShape::Box),
                        Some(b"cone") => Some(ChartBar3dShape::Cone),
                        Some(b"coneToMax") => Some(ChartBar3dShape::ConeToMax),
                        Some(b"cylinder") => Some(ChartBar3dShape::Cylinder),
                        Some(b"pyramid") => Some(ChartBar3dShape::Pyramid),
                        Some(b"pyramidToMax") => Some(ChartBar3dShape::PyramidToMax),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"sizeRepresents" => {
                    group.size_represents = match e.raw_attr(b"val")? {
                        Some(b"w") => Some(ChartSizeRepresents::Width),
                        Some(b"area") => Some(ChartSizeRepresents::Area),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"showNegBubbles" => {
                    group.show_negative_bubbles = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                b"splitType" => {
                    group.split_type = match e.raw_attr(b"val")? {
                        Some(b"auto") => Some(ChartOfPieSplitType::Auto),
                        Some(b"cust") => Some(ChartOfPieSplitType::Custom),
                        Some(b"percent") => Some(ChartOfPieSplitType::Percent),
                        Some(b"pos") => Some(ChartOfPieSplitType::Position),
                        Some(b"val") => Some(ChartOfPieSplitType::Value),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"splitPos" => {
                    group.split_position = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"custSplit" => {
                    parse_custom_split(xml, &e, &mut group.custom_split)?;
                }
                b"secondPieSize" => {
                    group.second_pie_size = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"dropLines" => {
                    group.drop_lines = Some(parse_chart_lines(xml, &e, theme_colors)?);
                }
                b"hiLowLines" => {
                    group.hi_low_lines = Some(parse_chart_lines(xml, &e, theme_colors)?);
                }
                b"serLines" => {
                    group.series_lines = Some(parse_chart_lines(xml, &e, theme_colors)?);
                }
                b"upDownBars" => {
                    group.up_down_bars = Some(parse_up_down_bars(xml, &e, theme_colors)?);
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("chart group")),
            _ => (),
        }
    }

    group.chart_type = resolve_chart_type(
        &group_tag,
        bar_dir.as_deref(),
        grouping.as_deref(),
        scatter_style.as_deref(),
        radar_style.as_deref(),
        of_pie_type.as_deref(),
        wireframe,
    );
    Ok(group)
}

/// Map a plot group element name and its layout attributes to a
/// [`ChartType`].
fn resolve_chart_type(
    group_tag: &[u8],
    bar_dir: Option<&str>,
    grouping: Option<&str>,
    scatter_style: Option<&str>,
    radar_style: Option<&str>,
    of_pie_type: Option<&str>,
    wireframe: bool,
) -> ChartType {
    match group_tag {
        b"barChart" | b"bar3DChart" => {
            let is_3d = group_tag == b"bar3DChart";
            let is_bar = bar_dir == Some("bar");
            match (is_bar, grouping, is_3d) {
                (true, Some("stacked"), false) => ChartType::BarStacked,
                (true, Some("percentStacked"), false) => ChartType::BarPercentStacked,
                (true, _, false) => ChartType::Bar,
                (true, Some("stacked"), true) => ChartType::Bar3DStacked,
                (true, Some("percentStacked"), true) => ChartType::Bar3DPercentStacked,
                (true, Some("standard"), true) => ChartType::Bar3DStandard,
                (true, _, true) => ChartType::Bar3D,
                (false, Some("stacked"), false) => ChartType::ColumnStacked,
                (false, Some("percentStacked"), false) => ChartType::ColumnPercentStacked,
                (false, _, false) => ChartType::Column,
                (false, Some("stacked"), true) => ChartType::Column3DStacked,
                (false, Some("percentStacked"), true) => ChartType::Column3DPercentStacked,
                (false, Some("standard"), true) => ChartType::Column3DStandard,
                (false, _, true) => ChartType::Column3D,
            }
        }
        b"lineChart" => match grouping {
            Some("stacked") => ChartType::LineStacked,
            Some("percentStacked") => ChartType::LinePercentStacked,
            _ => ChartType::Line,
        },
        b"line3DChart" => ChartType::Line3D,
        b"pieChart" => ChartType::Pie,
        b"pie3DChart" => ChartType::Pie3D,
        b"doughnutChart" => ChartType::Doughnut,
        b"ofPieChart" => match of_pie_type {
            Some("bar") => ChartType::BarOfPie,
            _ => ChartType::PieOfPie,
        },
        b"areaChart" | b"area3DChart" => {
            let is_3d = group_tag == b"area3DChart";
            match (grouping, is_3d) {
                (Some("stacked"), false) => ChartType::AreaStacked,
                (Some("percentStacked"), false) => ChartType::AreaPercentStacked,
                (_, false) => ChartType::Area,
                (Some("stacked"), true) => ChartType::Area3DStacked,
                (Some("percentStacked"), true) => ChartType::Area3DPercentStacked,
                (_, true) => ChartType::Area3D,
            }
        }
        b"scatterChart" => match scatter_style {
            Some("lineMarker") => ChartType::ScatterStraightWithMarkers,
            Some("line") => ChartType::ScatterStraight,
            Some("smoothMarker") => ChartType::ScatterSmoothWithMarkers,
            Some("smooth") => ChartType::ScatterSmooth,
            _ => ChartType::Scatter,
        },
        b"radarChart" => match radar_style {
            Some("filled") => ChartType::RadarFilled,
            Some("marker") => ChartType::RadarWithMarkers,
            _ => ChartType::Radar,
        },
        b"stockChart" => ChartType::Stock,
        b"bubbleChart" => ChartType::Bubble,
        b"surfaceChart" => {
            if wireframe {
                ChartType::ContourWireframe
            } else {
                ChartType::Contour
            }
        }
        b"surface3DChart" => {
            if wireframe {
                ChartType::Surface3DWireframe
            } else {
                ChartType::Surface3D
            }
        }
        _ => ChartType::Unknown,
    }
}

// ---------------------------------------------------------------------------
// Series
// ---------------------------------------------------------------------------

/// Parse a `c:ser` element.
fn parse_series<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartSeries, XlsxError> {
    let mut series = ChartSeries::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"idx" => {
                    series.index = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"order" => {
                    series.order = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"tx" => series.name = Some(parse_data_source(xml, &e)?),
                b"cat" | b"xVal" => series.categories = Some(parse_data_source(xml, &e)?),
                b"val" | b"yVal" => series.values = Some(parse_data_source(xml, &e)?),
                b"bubbleSize" => series.bubble_sizes = Some(parse_data_source(xml, &e)?),
                b"spPr" => series.format = parse_shape_properties(xml, &e, theme_colors)?,
                b"marker" => series.marker = Some(parse_marker(xml, &e, theme_colors)?),
                b"smooth" => {
                    series.smooth = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                b"invertIfNegative" => {
                    series.invert_if_negative = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                b"dPt" => {
                    if let Some(point) = parse_data_point(xml, &e, theme_colors)? {
                        series.points.push(point);
                    }
                }
                b"dLbls" => {
                    series.data_labels = Some(parse_data_labels(xml, &e, theme_colors)?);
                }
                b"trendline" => {
                    series
                        .trendlines
                        .push(parse_trendline(xml, &e, theme_colors)?);
                }
                b"errBars" => {
                    series
                        .error_bars
                        .push(parse_error_bars(xml, &e, theme_colors)?);
                }
                b"explosion" => {
                    series.explosion = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"bubble3D" => {
                    series.bubble_3d = val_attr_bool(&e)?;
                    skip_element(xml, &e)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("ser")),
            _ => (),
        }
    }
    Ok(series)
}

/// Parse a `c:dPt` (data point override) element.
fn parse_data_point<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<Option<ChartDataPoint>, XlsxError> {
    let mut index: Option<u32> = None;
    let mut format: Option<ChartFormat> = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"idx" => {
                    index = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"spPr" => format = parse_shape_properties(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("dPt")),
            _ => (),
        }
    }
    Ok(index.map(|index| ChartDataPoint { index, format }))
}

/// Parse a `c:marker` element inside a series.
fn parse_marker<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartMarker, XlsxError> {
    let mut marker = ChartMarker::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"symbol" => {
                    if let Some(symbol) = val_attr_string(xml, &e)? {
                        marker.marker_type = match symbol.as_str() {
                            "auto" => ChartMarkerType::Automatic,
                            "circle" => ChartMarkerType::Circle,
                            "dash" => ChartMarkerType::Dash,
                            "diamond" => ChartMarkerType::Diamond,
                            "dot" => ChartMarkerType::Dot,
                            "plus" => ChartMarkerType::Plus,
                            "square" => ChartMarkerType::Square,
                            "star" => ChartMarkerType::Star,
                            "triangle" => ChartMarkerType::Triangle,
                            "x" => ChartMarkerType::X,
                            "none" => ChartMarkerType::None,
                            _ => ChartMarkerType::Automatic,
                        };
                    }
                    skip_element(xml, &e)?;
                }
                b"size" => {
                    marker.size = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"spPr" => marker.format = parse_shape_properties(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("marker")),
            _ => (),
        }
    }
    Ok(marker)
}

// ---------------------------------------------------------------------------
// Data sources (c:tx / c:cat / c:val / c:xVal / c:yVal / c:bubbleSize)
// ---------------------------------------------------------------------------

/// Parse a series data reference wrapper containing a `strRef`, `numRef`,
/// `multiLvlStrRef`, `numLit` or `strLit`.
fn parse_data_source<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
) -> Result<ChartDataSource, XlsxError> {
    let mut source = ChartDataSource::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"strRef" | b"numRef" | b"multiLvlStrRef" => {
                    parse_data_reference(xml, &e, &mut source)?;
                }
                b"numLit" => parse_data_cache(xml, &e, &mut source, true)?,
                b"strLit" => parse_data_cache(xml, &e, &mut source, false)?,
                // Plain rich text series names (c:tx > c:rich) cache only.
                b"rich" => {
                    let (rich, _, _) = parse_rich_text(xml, &e, &[])?;
                    let text = rich.plain_text();
                    if !text.is_empty() {
                        source.values.push(Data::String(text));
                    }
                }
                b"v" => {
                    let text = read_text(xml, &e)?;
                    source.values.push(Data::String(text));
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("chart data source")),
            _ => (),
        }
    }
    Ok(source)
}

/// Parse a `strRef` / `numRef` / `multiLvlStrRef`: the formula plus the
/// cached values.
fn parse_data_reference<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    source: &mut ChartDataSource,
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"f" => source.formula = Some(read_text(xml, &e)?),
                b"numCache" => parse_data_cache(xml, &e, source, true)?,
                b"strCache" | b"multiLvlStrCache" => parse_data_cache(xml, &e, source, false)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("chart data reference")),
            _ => (),
        }
    }
    Ok(())
}

/// Parse a cache (`numCache`, `strCache`, `multiLvlStrCache`, `numLit` or
/// `strLit`) into the data source's values.
///
/// Multi-level caches keep all levels in [`ChartDataSource::levels`]
/// (innermost first, as in the document), with the innermost level also
/// mirrored into [`ChartDataSource::values`].
fn parse_data_cache<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    source: &mut ChartDataSource,
    numeric: bool,
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    let mut pt_count: Option<usize> = None;
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"formatCode" => source.number_format = Some(read_text(xml, &e)?),
                b"ptCount" => {
                    pt_count = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"lvl" => {
                    let level = parse_cache_level(xml, &e, numeric)?;
                    if source.levels.is_empty() {
                        source.values = level.clone();
                    }
                    source.levels.push(level);
                }
                b"pt" => {
                    let idx: Option<usize> = e
                        .raw_attr(b"idx")?
                        .and_then(parse_bytes);
                    let value = parse_cache_point(xml, &e, numeric)?;
                    let idx = idx.unwrap_or(source.values.len());
                    if source.values.len() <= idx {
                        source.values.resize(idx + 1, Data::Empty);
                    }
                    source.values[idx] = value;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("chart data cache")),
            _ => (),
        }
    }
    if let Some(count) = pt_count {
        if source.values.len() < count {
            source.values.resize(count, Data::Empty);
        }
        for level in &mut source.levels {
            if level.len() < count {
                level.resize(count, Data::Empty);
            }
        }
    }
    Ok(())
}

/// Parse one `c:lvl` of a `c:multiLvlStrCache` into a vector of points.
fn parse_cache_level<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    numeric: bool,
) -> Result<Vec<Data>, XlsxError> {
    let mut points = Vec::new();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"pt" => {
                    let idx: Option<usize> = e.raw_attr(b"idx")?.and_then(parse_bytes);
                    let value = parse_cache_point(xml, &e, numeric)?;
                    let idx = idx.unwrap_or(points.len());
                    if points.len() <= idx {
                        points.resize(idx + 1, Data::Empty);
                    }
                    points[idx] = value;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("lvl")),
            _ => (),
        }
    }
    Ok(points)
}

/// Parse one `c:pt` element's `c:v` child into a [`Data`] value.
fn parse_cache_point<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    numeric: bool,
) -> Result<Data, XlsxError> {
    let mut value = Data::Empty;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => {
                if e.local_name().as_ref() == b"v" {
                    let text = read_text(xml, &e)?;
                    value = if numeric {
                        match text.trim().parse::<f64>() {
                            Ok(v) => Data::Float(v),
                            Err(_) => Data::String(text),
                        }
                    } else {
                        Data::String(text)
                    };
                } else {
                    skip_element(xml, &e)?;
                }
            }
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("pt")),
            _ => (),
        }
    }
    Ok(value)
}

// ---------------------------------------------------------------------------
// Title, legend, view3D
// ---------------------------------------------------------------------------

/// Parse a `c:title` element (chart or axis title).
fn parse_title<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartTitle, XlsxError> {
    let mut title = ChartTitle::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"tx" => (),
                b"rich" => {
                    let (rich, font, rotation) = parse_rich_text(xml, &e, theme_colors)?;
                    title.rich = Some(rich);
                    if title.font.is_none() {
                        title.font = font;
                    }
                    if title.text_rotation.is_none() {
                        title.text_rotation = rotation;
                    }
                }
                b"strRef" => {
                    let mut source = ChartDataSource::default();
                    parse_data_reference(xml, &e, &mut source)?;
                    title.formula = source.formula;
                    title.cached = source.values.into_iter().find_map(|v| match v {
                        Data::String(s) => Some(s),
                        _ => None,
                    });
                }
                b"overlay" => {
                    title.overlay = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"txPr" => {
                    let (font, rotation) = parse_text_properties(xml, &e, theme_colors)?;
                    if title.font.is_none() {
                        title.font = font;
                    }
                    if title.text_rotation.is_none() {
                        title.text_rotation = rotation;
                    }
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("title")),
            _ => (),
        }
    }
    Ok(title)
}

/// Parse a `c:legend` element.
fn parse_legend<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartLegend, XlsxError> {
    let mut legend = ChartLegend::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"legendPos" => {
                    if let Some(pos) = val_attr_string(xml, &e)? {
                        legend.position = match pos.as_str() {
                            "l" => ChartLegendPosition::Left,
                            "t" => ChartLegendPosition::Top,
                            "b" => ChartLegendPosition::Bottom,
                            "tr" => ChartLegendPosition::TopRight,
                            _ => ChartLegendPosition::Right,
                        };
                    }
                    skip_element(xml, &e)?;
                }
                b"overlay" => {
                    legend.overlay = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"txPr" => legend.font = parse_text_properties(xml, &e, theme_colors)?.0,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("legend")),
            _ => (),
        }
    }
    Ok(legend)
}

/// Parse a `c:view3D` element.
fn parse_view_3d<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
) -> Result<ChartView3d, XlsxError> {
    let mut view = ChartView3d::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => {
                match e.local_name().as_ref() {
                    b"rotX" => view.rot_x = val_attr(&e)?,
                    b"rotY" => view.rot_y = val_attr(&e)?,
                    // The file stores twice the perspective angle.
                    b"perspective" => view.perspective = val_attr::<u32>(&e)?.map(|v| v / 2),
                    b"depthPercent" => view.depth_percent = val_attr(&e)?,
                    b"hPercent" => view.height_percent = val_attr(&e)?,
                    b"rAngAx" => view.right_angle_axes = val_attr_bool(&e)?,
                    _ => (),
                }
                skip_element(xml, &e)?;
            }
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("view3D")),
            _ => (),
        }
    }
    Ok(view)
}

// ---------------------------------------------------------------------------
// Axes
// ---------------------------------------------------------------------------

/// Parse a `c:catAx` / `c:valAx` / `c:dateAx` / `c:serAx` element.
fn parse_axis<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    axis_type: ChartAxisType,
    theme_colors: &[Color],
) -> Result<ChartAxis, XlsxError> {
    let mut axis = ChartAxis::new(axis_type);
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"axId" => {
                    axis.id = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"scaling" => parse_axis_scaling(xml, &e, &mut axis)?,
                b"delete" => {
                    axis.hidden = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"axPos" => {
                    if let Some(pos) = val_attr_string(xml, &e)? {
                        axis.position = match pos.as_str() {
                            "b" => Some(ChartAxisPosition::Bottom),
                            "l" => Some(ChartAxisPosition::Left),
                            "r" => Some(ChartAxisPosition::Right),
                            "t" => Some(ChartAxisPosition::Top),
                            _ => None,
                        };
                    }
                    skip_element(xml, &e)?;
                }
                b"majorGridlines" => {
                    axis.major_gridlines = true;
                    skip_element(xml, &e)?;
                }
                b"minorGridlines" => {
                    axis.minor_gridlines = true;
                    skip_element(xml, &e)?;
                }
                b"title" => axis.title = Some(parse_title(xml, &e, theme_colors)?),
                b"numFmt" => {
                    if let Some(code) = e.raw_attr(b"formatCode")? {
                        axis.number_format = Some(decode_attr(&xml.decoder(), code)?);
                    }
                    skip_element(xml, &e)?;
                }
                b"majorUnit" => {
                    axis.major_unit = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"minorUnit" => {
                    axis.minor_unit = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"majorTickMark" => {
                    axis.major_tick_mark = parse_tick_mark(&e)?;
                    skip_element(xml, &e)?;
                }
                b"minorTickMark" => {
                    axis.minor_tick_mark = parse_tick_mark(&e)?;
                    skip_element(xml, &e)?;
                }
                b"tickLblPos" => {
                    axis.tick_label_position = match e.raw_attr(b"val")? {
                        Some(b"nextTo") => Some(ChartTickLabelPosition::NextTo),
                        Some(b"high") => Some(ChartTickLabelPosition::High),
                        Some(b"low") => Some(ChartTickLabelPosition::Low),
                        Some(b"none") => Some(ChartTickLabelPosition::None),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"crosses" => {
                    axis.crosses = match e.raw_attr(b"val")? {
                        Some(b"autoZero") => Some(ChartAxisCrosses::AutoZero),
                        Some(b"min") => Some(ChartAxisCrosses::Min),
                        Some(b"max") => Some(ChartAxisCrosses::Max),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"crossesAt" => {
                    axis.crosses_at = val_attr(&e)?;
                    if axis.crosses_at.is_some() {
                        axis.crosses = Some(ChartAxisCrosses::At);
                    }
                    skip_element(xml, &e)?;
                }
                b"crossAx" => {
                    axis.crosses_axis_id = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"crossBetween" => {
                    axis.cross_between = match e.raw_attr(b"val")? {
                        Some(b"between") => Some(ChartCrossBetween::Between),
                        Some(b"midCat") => Some(ChartCrossBetween::MidCat),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"dispUnits" => parse_display_units(xml, &e, &mut axis)?,
                b"tickLblSkip" => {
                    axis.tick_label_skip = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"tickMarkSkip" => {
                    axis.tick_mark_skip = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"lblOffset" => {
                    axis.label_offset = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"spPr" => axis.format = parse_shape_properties(xml, &e, theme_colors)?,
                b"txPr" => {
                    let (font, rotation) = parse_text_properties(xml, &e, theme_colors)?;
                    axis.font = font;
                    axis.text_rotation = rotation;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("axis")),
            _ => (),
        }
    }
    Ok(axis)
}

/// Parse the `c:scaling` element of an axis.
fn parse_axis_scaling<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    axis: &mut ChartAxis,
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => {
                match e.local_name().as_ref() {
                    b"orientation" => {
                        axis.reverse = matches!(e.raw_attr(b"val")?, Some(b"maxMin"));
                    }
                    b"min" => axis.min = val_attr(&e)?,
                    b"max" => axis.max = val_attr(&e)?,
                    b"logBase" => axis.log_base = val_attr(&e)?,
                    _ => (),
                }
                skip_element(xml, &e)?;
            }
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("scaling")),
            _ => (),
        }
    }
    Ok(())
}

fn parse_tick_mark(e: &BytesStart) -> Result<Option<ChartTickMark>, XlsxError> {
    Ok(match e.raw_attr(b"val")? {
        Some(b"none") => Some(ChartTickMark::None),
        Some(b"in") => Some(ChartTickMark::Inside),
        Some(b"out") => Some(ChartTickMark::Outside),
        Some(b"cross") => Some(ChartTickMark::Cross),
        _ => None,
    })
}

/// Parse a `c:dispUnits` element into the axis display units.
fn parse_display_units<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    axis: &mut ChartAxis,
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"builtInUnit" => {
                    axis.display_units = match e.raw_attr(b"val")? {
                        Some(b"hundreds") => Some(ChartDisplayUnits::Hundreds),
                        Some(b"thousands") => Some(ChartDisplayUnits::Thousands),
                        Some(b"tenThousands") => Some(ChartDisplayUnits::TenThousands),
                        Some(b"hundredThousands") => Some(ChartDisplayUnits::HundredThousands),
                        Some(b"millions") => Some(ChartDisplayUnits::Millions),
                        Some(b"tenMillions") => Some(ChartDisplayUnits::TenMillions),
                        Some(b"hundredMillions") => Some(ChartDisplayUnits::HundredMillions),
                        Some(b"billions") => Some(ChartDisplayUnits::Billions),
                        Some(b"trillions") => Some(ChartDisplayUnits::Trillions),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"custUnit" => {
                    axis.display_units = val_attr::<f64>(&e)?.map(ChartDisplayUnits::Custom);
                    skip_element(xml, &e)?;
                }
                b"dispUnitsLbl" => {
                    axis.display_units_label = true;
                    skip_element(xml, &e)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("dispUnits")),
            _ => (),
        }
    }
    Ok(())
}

/// Parse a `c:dLbls` (data labels) element.
fn parse_data_labels<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartDataLabels, XlsxError> {
    let mut labels = ChartDataLabels::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"showVal" => {
                    labels.show_value = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"showCatName" => {
                    labels.show_category_name = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"showSerName" => {
                    labels.show_series_name = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"showLegendKey" => {
                    labels.show_legend_key = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"showPercent" => {
                    labels.show_percent = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"showBubbleSize" => {
                    labels.show_bubble_size = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"dLblPos" => {
                    labels.position = data_label_position(&e)?;
                    skip_element(xml, &e)?;
                }
                b"numFmt" => {
                    if let Some(code) = e.raw_attr(b"formatCode")? {
                        labels.number_format = Some(decode_attr(&xml.decoder(), code)?);
                    }
                    skip_element(xml, &e)?;
                }
                b"spPr" => labels.format = parse_shape_properties(xml, &e, theme_colors)?,
                b"txPr" => {
                    let (font, rotation) = parse_text_properties(xml, &e, theme_colors)?;
                    labels.font = font;
                    labels.text_rotation = rotation;
                }
                b"dLbl" => {
                    labels
                        .point_labels
                        .push(parse_data_label(xml, &e, theme_colors)?);
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("dLbls")),
            _ => (),
        }
    }
    Ok(labels)
}

/// Parse the `c:secondPiePt` children of `c:custSplit`.
fn parse_custom_split<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    points: &mut Vec<u32>,
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"secondPiePt" => {
                    if let Some(idx) = val_attr(&e)? {
                        points.push(idx);
                    }
                    skip_element(xml, &e)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("custSplit")),
            _ => (),
        }
    }
    Ok(())
}

/// Parse a `c:dLblPos` element's `val` attribute.
fn data_label_position(e: &BytesStart) -> Result<Option<ChartDataLabelPosition>, XlsxError> {
    Ok(match e.raw_attr(b"val")? {
        Some(b"ctr") => Some(ChartDataLabelPosition::Center),
        Some(b"inEnd") => Some(ChartDataLabelPosition::InsideEnd),
        Some(b"inBase") => Some(ChartDataLabelPosition::InsideBase),
        Some(b"outEnd") => Some(ChartDataLabelPosition::OutsideEnd),
        Some(b"l") => Some(ChartDataLabelPosition::Left),
        Some(b"r") => Some(ChartDataLabelPosition::Right),
        Some(b"t") => Some(ChartDataLabelPosition::Above),
        Some(b"b") => Some(ChartDataLabelPosition::Below),
        Some(b"bestFit") => Some(ChartDataLabelPosition::BestFit),
        _ => None,
    })
}

/// Parse a `c:dLbl` (per-point data label override) element.
fn parse_data_label<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartDataLabel, XlsxError> {
    let mut label = ChartDataLabel::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"idx" => {
                    label.index = val_attr(&e)?.unwrap_or(0);
                    skip_element(xml, &e)?;
                }
                b"delete" => {
                    label.delete = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"tx" => (),
                b"rich" => {
                    let (rich, font, _) = parse_rich_text(xml, &e, theme_colors)?;
                    let text = rich.plain_text();
                    if !text.is_empty() {
                        label.text = Some(text);
                    }
                    if label.font.is_none() {
                        label.font = font;
                    }
                }
                b"dLblPos" => {
                    label.position = data_label_position(&e)?;
                    skip_element(xml, &e)?;
                }
                b"numFmt" => {
                    if let Some(code) = e.raw_attr(b"formatCode")? {
                        label.number_format = Some(decode_attr(&xml.decoder(), code)?);
                    }
                    skip_element(xml, &e)?;
                }
                b"spPr" => label.format = parse_shape_properties(xml, &e, theme_colors)?,
                b"txPr" => {
                    let (font, _) = parse_text_properties(xml, &e, theme_colors)?;
                    if label.font.is_none() {
                        label.font = font;
                    }
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("dLbl")),
            _ => (),
        }
    }
    Ok(label)
}

/// Parse a `c:trendline` element.
fn parse_trendline<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartTrendline, XlsxError> {
    let mut trendline = ChartTrendline::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"name" => trendline.name = Some(read_text(xml, &e)?),
                b"trendlineType" => {
                    trendline.trendline_type = match e.raw_attr(b"val")? {
                        Some(b"exp") => ChartTrendlineType::Exponential,
                        Some(b"log") => ChartTrendlineType::Logarithmic,
                        Some(b"movingAvg") => ChartTrendlineType::MovingAverage,
                        Some(b"poly") => ChartTrendlineType::Polynomial,
                        Some(b"power") => ChartTrendlineType::Power,
                        _ => ChartTrendlineType::Linear,
                    };
                    skip_element(xml, &e)?;
                }
                b"order" => {
                    trendline.order = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"period" => {
                    trendline.period = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"forward" => {
                    trendline.forward = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"backward" => {
                    trendline.backward = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"intercept" => {
                    trendline.intercept = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"dispEq" => {
                    trendline.display_equation = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"dispRSqr" => {
                    trendline.display_r_squared = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"spPr" => trendline.format = parse_shape_properties(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("trendline")),
            _ => (),
        }
    }
    Ok(trendline)
}

/// Parse a `c:errBars` element.
fn parse_error_bars<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartErrorBars, XlsxError> {
    let mut bars = ChartErrorBars::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"errDir" => {
                    bars.direction = match e.raw_attr(b"val")? {
                        Some(b"x") => Some(ChartErrorBarsDirection::X),
                        Some(b"y") => Some(ChartErrorBarsDirection::Y),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"errBarType" => {
                    bars.error_type = match e.raw_attr(b"val")? {
                        Some(b"minus") => ChartErrorBarsType::Minus,
                        Some(b"plus") => ChartErrorBarsType::Plus,
                        _ => ChartErrorBarsType::Both,
                    };
                    skip_element(xml, &e)?;
                }
                b"errValType" => {
                    bars.value_type = match e.raw_attr(b"val")? {
                        Some(b"cust") => ChartErrorBarsValueType::Custom,
                        Some(b"percentage") => ChartErrorBarsValueType::Percentage,
                        Some(b"stdDev") => ChartErrorBarsValueType::StandardDeviation,
                        Some(b"stdErr") => ChartErrorBarsValueType::StandardError,
                        _ => ChartErrorBarsValueType::FixedValue,
                    };
                    skip_element(xml, &e)?;
                }
                b"noEndCap" => {
                    bars.no_end_cap = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"val" => {
                    bars.value = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"plus" => bars.plus_values = Some(parse_data_source(xml, &e)?),
                b"minus" => bars.minus_values = Some(parse_data_source(xml, &e)?),
                b"spPr" => bars.format = parse_shape_properties(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("errBars")),
            _ => (),
        }
    }
    Ok(bars)
}

/// Parse a `c:dropLines` / `c:hiLowLines` / `c:serLines` element.
fn parse_chart_lines<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartLines, XlsxError> {
    let mut lines = ChartLines::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"spPr" => lines.format = parse_shape_properties(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("chart lines")),
            _ => (),
        }
    }
    Ok(lines)
}

/// Parse a `c:upDownBars` element.
fn parse_up_down_bars<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartUpDownBars, XlsxError> {
    let mut bars = ChartUpDownBars::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"gapWidth" => {
                    bars.gap_width = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"upBars" => bars.up_format = parse_bars_format(xml, &e, theme_colors)?,
                b"downBars" => bars.down_format = parse_bars_format(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("upDownBars")),
            _ => (),
        }
    }
    Ok(bars)
}

/// Parse the `c:spPr` inside a `c:upBars` / `c:downBars` element.
fn parse_bars_format<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<Option<ChartFormat>, XlsxError> {
    let mut format = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"spPr" => format = parse_shape_properties(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("upBars")),
            _ => (),
        }
    }
    Ok(format)
}

/// Parse a `c:dTable` (data table) element.
fn parse_data_table<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartDataTable, XlsxError> {
    let mut table = ChartDataTable::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"showHorzBorder" => {
                    table.show_horizontal_border = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"showVertBorder" => {
                    table.show_vertical_border = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"showOutline" => {
                    table.show_outline = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"showKeys" => {
                    table.show_legend_keys = val_attr_bool(&e)?.unwrap_or(false);
                    skip_element(xml, &e)?;
                }
                b"txPr" => table.font = parse_text_properties(xml, &e, theme_colors)?.0,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("dTable")),
            _ => (),
        }
    }
    Ok(table)
}

// ---------------------------------------------------------------------------
// Rich text and fonts
// ---------------------------------------------------------------------------

/// Parse the `rot` attribute of an `a:bodyPr` element into degrees.
fn body_rotation(e: &BytesStart) -> Result<Option<f64>, XlsxError> {
    // Stored in 1/60000ths of a degree.
    Ok(e.raw_attr(b"rot")?
        .and_then(parse_bytes::<f64>)
        .map(|r| r / 60000.0))
}

/// Parse a `c:rich` element into a [`RichText`] plus the paragraph default
/// font and the body text rotation, if any.
fn parse_rich_text<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<(RichText, Option<Font>, Option<f64>), XlsxError> {
    let mut rich = RichText::new();
    let mut default_font: Option<Font> = None;
    let mut rotation: Option<f64> = None;
    let mut paragraphs = 0usize;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"bodyPr" => {
                    rotation = body_rotation(&e)?;
                    skip_element(xml, &e)?;
                }
                b"p" => {
                    paragraphs += 1;
                    if paragraphs > 1 && !rich.is_empty() {
                        rich.push_text("\n".to_string());
                    }
                }
                b"pPr" => {
                    if let Some(font) = parse_paragraph_properties(xml, &e, theme_colors)? {
                        default_font.get_or_insert(font);
                    }
                }
                b"r" => {
                    if let Some(run) = parse_text_run(xml, &e, theme_colors)? {
                        rich.push(run);
                    }
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("rich")),
            _ => (),
        }
    }
    Ok((rich, default_font, rotation))
}

/// Parse an `a:pPr` element, returning the `a:defRPr` font if present.
fn parse_paragraph_properties<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<Option<Font>, XlsxError> {
    let mut font = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => {
                if e.local_name().as_ref() == b"defRPr" {
                    font = parse_run_properties(xml, &e, theme_colors)?;
                } else {
                    skip_element(xml, &e)?;
                }
            }
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("pPr")),
            _ => (),
        }
    }
    Ok(font)
}

/// Parse an `a:r` (text run) element.
fn parse_text_run<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<Option<TextRun>, XlsxError> {
    let mut font: Option<Font> = None;
    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"rPr" => font = parse_run_properties(xml, &e, theme_colors)?,
                b"t" => text.push_str(&read_text(xml, &e)?),
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("r")),
            _ => (),
        }
    }
    if text.is_empty() && font.is_none() {
        return Ok(None);
    }
    Ok(Some(TextRun { text, font }))
}

/// Parse an `a:rPr` / `a:defRPr` element into a [`Font`].
///
/// Returns `None` if no font properties are set.
fn parse_run_properties<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<Option<Font>, XlsxError> {
    let mut font = Font::new();
    let mut any = false;

    for attr in parent.iter_raw_attrs() {
        let (key, value) = attr.map_err(XlsxError::XmlAttr)?;
        match key {
            // Font size is stored in hundredths of a point.
            b"sz" => {
                if let Some(sz) = parse_bytes::<f64>(value) {
                    font.size = Some(sz / 100.0);
                    any = true;
                }
            }
            b"b" => {
                if matches!(value, b"1" | b"true") {
                    font.weight = FontWeight::Bold;
                    any = true;
                }
            }
            b"i" => {
                if matches!(value, b"1" | b"true") {
                    font.style = FontStyle::Italic;
                    any = true;
                }
            }
            b"u" => {
                font.underline = match value {
                    b"sng" => UnderlineStyle::Single,
                    b"dbl" => UnderlineStyle::Double,
                    _ => UnderlineStyle::None,
                };
                if font.underline != UnderlineStyle::None {
                    any = true;
                }
            }
            b"strike" => {
                if matches!(value, b"sngStrike" | b"dblStrike") {
                    font.strikethrough = true;
                    any = true;
                }
            }
            _ => (),
        }
    }

    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"solidFill" => {
                    if let Some(color) = parse_color_container(xml, &e, theme_colors)? {
                        font.color = Some(color);
                        any = true;
                    }
                }
                b"latin" => {
                    if let Some(typeface) = e.raw_attr(b"typeface")? {
                        let name = decode_attr(&xml.decoder(), typeface)?;
                        // Skip theme placeholders like "+mn-lt".
                        if !name.starts_with('+') {
                            font.name = Some(name);
                            any = true;
                        }
                    }
                    skip_element(xml, &e)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("rPr")),
            _ => (),
        }
    }

    Ok(any.then_some(font))
}

/// Parse a `c:txPr` element, returning the default font of the first
/// paragraph and the body text rotation, if any.
fn parse_text_properties<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<(Option<Font>, Option<f64>), XlsxError> {
    let mut font: Option<Font> = None;
    let mut rotation: Option<f64> = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"bodyPr" => {
                    rotation = body_rotation(&e)?;
                    skip_element(xml, &e)?;
                }
                b"p" => (),
                b"pPr" => {
                    if let Some(f) = parse_paragraph_properties(xml, &e, theme_colors)? {
                        font.get_or_insert(f);
                    }
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("txPr")),
            _ => (),
        }
    }
    Ok((font, rotation))
}

// ---------------------------------------------------------------------------
// Shape properties (fills and lines)
// ---------------------------------------------------------------------------

/// Parse an `a:spPr` (shape properties) element into a [`ChartFormat`].
///
/// Returns `None` if neither a fill nor a line is specified.
fn parse_shape_properties<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<Option<ChartFormat>, XlsxError> {
    let mut format = ChartFormat::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"noFill" => {
                    format.fill = Some(ChartFill::None);
                    skip_element(xml, &e)?;
                }
                b"solidFill" => {
                    if let Some(color) = parse_color_container(xml, &e, theme_colors)? {
                        format.fill = Some(ChartFill::Solid(color));
                    }
                }
                b"gradFill" => {
                    let (stops, angle) = parse_gradient_fill(xml, &e, theme_colors)?;
                    format.fill = Some(ChartFill::Gradient(stops));
                    format.gradient_angle = angle;
                }
                b"pattFill" => {
                    format.fill = Some(parse_pattern_fill(xml, &e, theme_colors)?);
                }
                b"ln" => format.line = Some(parse_line(xml, &e, theme_colors)?),
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("spPr")),
            _ => (),
        }
    }
    if format.fill.is_none() && format.line.is_none() {
        Ok(None)
    } else {
        Ok(Some(format))
    }
}

/// Parse an `a:ln` (line properties) element.
fn parse_line<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartLine, XlsxError> {
    let mut line = ChartLine::default();

    // Line width is stored in EMUs (12700 per point).
    if let Some(w) = parent.raw_attr(b"w")? {
        if let Some(w) = parse_bytes::<f64>(w) {
            line.width = Some(w / 12700.0);
        }
    }

    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"noFill" => {
                    line.hidden = true;
                    skip_element(xml, &e)?;
                }
                b"solidFill" => {
                    line.color = parse_color_container(xml, &e, theme_colors)?;
                }
                b"prstDash" => {
                    if let Some(dash) = val_attr_string(xml, &e)? {
                        line.dash_type = Some(match dash.as_str() {
                            "dot" => ChartLineDashType::Dot,
                            "dash" => ChartLineDashType::Dash,
                            "dashDot" => ChartLineDashType::DashDot,
                            "lgDash" => ChartLineDashType::LongDash,
                            "lgDashDot" => ChartLineDashType::LongDashDot,
                            "lgDashDotDot" => ChartLineDashType::LongDashDotDot,
                            "sysDash" => ChartLineDashType::SystemDash,
                            "sysDot" => ChartLineDashType::SystemDot,
                            "sysDashDot" => ChartLineDashType::SystemDashDot,
                            "sysDashDotDot" => ChartLineDashType::SystemDashDotDot,
                            _ => ChartLineDashType::Solid,
                        });
                    }
                    skip_element(xml, &e)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("ln")),
            _ => (),
        }
    }
    Ok(line)
}

/// Parse an `a:gradFill` element into its gradient stops and the linear
/// gradient angle in degrees, if any.
fn parse_gradient_fill<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<(Vec<ChartGradientStop>, Option<f64>), XlsxError> {
    let mut stops = Vec::new();
    let mut angle: Option<f64> = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"gsLst" => (),
                b"gs" => {
                    // Stop position is stored in thousandths of a percent.
                    let position = e
                        .raw_attr(b"pos")?
                        .and_then(parse_bytes::<f64>)
                        .map_or(0.0, |p| p / 1000.0);
                    if let Some(color) = parse_color_container(xml, &e, theme_colors)? {
                        stops.push(ChartGradientStop { position, color });
                    }
                }
                b"lin" => {
                    // Angle is stored in 1/60000ths of a degree.
                    angle = e
                        .raw_attr(b"ang")?
                        .and_then(parse_bytes::<f64>)
                        .map(|a| a / 60000.0);
                    skip_element(xml, &e)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("gradFill")),
            _ => (),
        }
    }
    Ok((stops, angle))
}

/// Parse an `a:pattFill` element.
fn parse_pattern_fill<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartFill, XlsxError> {
    let pattern = match parent.raw_attr(b"prst")? {
        Some(p) => decode_attr(&xml.decoder(), p)?,
        None => String::new(),
    };
    let mut foreground = None;
    let mut background = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"fgClr" => foreground = parse_color_container(xml, &e, theme_colors)?,
                b"bgClr" => background = parse_color_container(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("pattFill")),
            _ => (),
        }
    }
    Ok(ChartFill::Pattern {
        pattern,
        foreground,
        background,
    })
}

// ---------------------------------------------------------------------------
// DrawingML colors
// ---------------------------------------------------------------------------

/// Parse the first color child (`a:srgbClr`, `a:schemeClr`, `a:sysClr` or
/// `a:prstClr`) of a container element such as `a:solidFill` or `a:gs`, and
/// consume the container.
fn parse_color_container<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<Option<Color>, XlsxError> {
    let mut color: Option<Color> = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"srgbClr" => {
                    let base = e
                        .raw_attr(b"val")?
                        .and_then(parse_hex_color);
                    let parsed = parse_color_modifiers(xml, &e, base)?;
                    color = color.or(parsed);
                }
                b"schemeClr" => {
                    let base = match e.raw_attr(b"val")? {
                        Some(name) => scheme_color(name, theme_colors),
                        None => None,
                    };
                    let parsed = parse_color_modifiers(xml, &e, base)?;
                    color = color.or(parsed);
                }
                b"sysClr" => {
                    let base = match e.raw_attr(b"lastClr")? {
                        Some(v) => parse_hex_color(v),
                        None => match e.raw_attr(b"val")? {
                            Some(b"window") => Some(Color::rgb(255, 255, 255)),
                            Some(b"windowText") => Some(Color::rgb(0, 0, 0)),
                            _ => None,
                        },
                    };
                    let parsed = parse_color_modifiers(xml, &e, base)?;
                    color = color.or(parsed);
                }
                b"prstClr" => {
                    let base = e.raw_attr(b"val")?.and_then(preset_color);
                    let parsed = parse_color_modifiers(xml, &e, base)?;
                    color = color.or(parsed);
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("color")),
            _ => (),
        }
    }
    Ok(color)
}

/// Consume the children of a color element, applying the common color
/// modifiers (`alpha`, `lumMod`, `lumOff`, `shade`, `tint`) to the base
/// color.
///
/// The luminance modifiers are approximated in RGB space, which matches the
/// common approach of other readers.
fn parse_color_modifiers<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    base: Option<Color>,
) -> Result<Option<Color>, XlsxError> {
    let mut color = base;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let modifier = e.local_name().as_ref().to_vec();
                let value = val_attr::<f64>(&e)?.map(|v| v / 100_000.0);
                if let (Some(c), Some(v)) = (color, value) {
                    color = Some(apply_color_modifier(c, &modifier, v));
                }
                skip_element(xml, &e)?;
            }
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("color modifiers")),
            _ => (),
        }
    }
    Ok(color)
}

fn apply_color_modifier(color: Color, modifier: &[u8], value: f64) -> Color {
    let clamp = |v: f64| v.round().clamp(0.0, 255.0) as u8;
    let per_channel = |color: Color, f: &dyn Fn(f64) -> f64| {
        Color::new(
            color.alpha,
            clamp(f(color.red as f64)),
            clamp(f(color.green as f64)),
            clamp(f(color.blue as f64)),
        )
    };
    match modifier {
        b"alpha" => Color::new(clamp(value * 255.0), color.red, color.green, color.blue),
        b"lumMod" | b"shade" => per_channel(color, &|c| c * value),
        b"lumOff" => per_channel(color, &|c| c + 255.0 * value),
        b"tint" => per_channel(color, &|c| c * value + 255.0 * (1.0 - value)),
        _ => color,
    }
}

/// Parse a 6-character hex color value such as `FF9900`.
fn parse_hex_color(bytes: &[u8]) -> Option<Color> {
    let s = std::str::from_utf8(bytes).ok()?;
    let s = s.trim();
    if s.len() != 6 {
        return None;
    }
    let argb = u32::from_str_radix(s, 16).ok()?;
    Some(Color::from_argb(0xFF00_0000 | argb))
}

/// Resolve a scheme color name (e.g. `accent1`, `tx1`, `bg2`) against the
/// workbook theme palette.
///
/// The palette is ordered `lt1, dk1, lt2, dk2, accent1..accent6, hlink,
/// folHlink` (Excel's theme indexing).
fn scheme_color(name: &[u8], theme_colors: &[Color]) -> Option<Color> {
    let index: usize = match name {
        b"bg1" | b"lt1" => 0,
        b"tx1" | b"dk1" => 1,
        b"bg2" | b"lt2" => 2,
        b"tx2" | b"dk2" => 3,
        b"accent1" => 4,
        b"accent2" => 5,
        b"accent3" => 6,
        b"accent4" => 7,
        b"accent5" => 8,
        b"accent6" => 9,
        b"hlink" => 10,
        b"folHlink" => 11,
        _ => return None,
    };
    theme_colors.get(index).copied()
}

/// Resolve a small set of common preset color names (`a:prstClr`).
fn preset_color(name: &[u8]) -> Option<Color> {
    let (r, g, b) = match name {
        b"black" => (0, 0, 0),
        b"white" => (255, 255, 255),
        b"red" => (255, 0, 0),
        b"green" => (0, 128, 0),
        b"lime" => (0, 255, 0),
        b"blue" => (0, 0, 255),
        b"yellow" => (255, 255, 0),
        b"cyan" | b"aqua" => (0, 255, 255),
        b"magenta" | b"fuchsia" => (255, 0, 255),
        b"gray" | b"grey" => (128, 128, 128),
        b"orange" => (255, 165, 0),
        b"purple" => (128, 0, 128),
        _ => return None,
    };
    Some(Color::rgb(r, g, b))
}

// ---------------------------------------------------------------------------
// Chart-ex parts (Excel 2016+ funnel/treemap/sunburst/histogram/pareto/
// box & whisker/waterfall/filled map charts, `cx:` namespace)
// ---------------------------------------------------------------------------

/// Parse a `cx:chartSpace` document into a [`Chart`].
pub(crate) fn parse_chartex_space<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    theme_colors: &[Color],
) -> Result<Chart, XlsxError> {
    let mut chart = Chart::default();
    // Data id -> (categories, values), from cx:chartData.
    let mut data: std::collections::HashMap<String, (Option<ChartDataSource>, Option<ChartDataSource>)> =
        std::collections::HashMap::new();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"chartSpace" | b"chart" | b"plotArea" | b"plotAreaRegion" => (),
                b"chartData" => parse_chartex_data(xml, &e, &mut data)?,
                b"title" => chart.title = Some(parse_chartex_title(xml, &e, theme_colors)?),
                b"series" => {
                    let group = parse_chartex_series(xml, &e, &data, theme_colors)?;
                    chart.groups.push(group);
                }
                b"axis" => chart.axes.push(parse_chartex_axis(xml, &e, theme_colors)?),
                b"legend" => {
                    let mut legend = ChartLegend::default();
                    if let Some(pos) = e.raw_attr(b"pos")? {
                        legend.position = match pos {
                            b"l" => ChartLegendPosition::Left,
                            b"t" => ChartLegendPosition::Top,
                            b"b" => ChartLegendPosition::Bottom,
                            b"tr" => ChartLegendPosition::TopRight,
                            _ => ChartLegendPosition::Right,
                        };
                    }
                    if let Some(overlay) = e.raw_attr(b"overlay")? {
                        legend.overlay = matches!(overlay, b"1" | b"true");
                    }
                    chart.legend = Some(legend);
                    skip_element(xml, &e)?;
                }
                b"spPr" => chart.format = parse_shape_properties(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::Eof => break,
            _ => (),
        }
    }

    // A pareto chart is a histogram (clusteredColumn) plus a paretoLine
    // series; report the whole chart as Pareto in that case.
    if chart
        .groups
        .iter()
        .any(|g| g.chart_type == ChartType::Pareto)
    {
        for group in &mut chart.groups {
            if group.chart_type == ChartType::Histogram {
                group.chart_type = ChartType::Pareto;
            }
        }
    }
    Ok(chart)
}

/// Parse the `cx:chartData` element: literal data dimensions keyed by id.
fn parse_chartex_data<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    data: &mut std::collections::HashMap<String, (Option<ChartDataSource>, Option<ChartDataSource>)>,
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"data" => {
                    let id = match e.raw_attr(b"id")? {
                        Some(id) => decode_attr(&xml.decoder(), id)?,
                        None => String::new(),
                    };
                    let entry = parse_chartex_data_entry(xml, &e)?;
                    data.insert(id, entry);
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("chartData")),
            _ => (),
        }
    }
    Ok(())
}

/// Parse one `cx:data` element into its category and value dimensions.
fn parse_chartex_data_entry<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
) -> Result<(Option<ChartDataSource>, Option<ChartDataSource>), XlsxError> {
    let mut categories = None;
    let mut values = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"strDim" => categories = Some(parse_chartex_dimension(xml, &e, false)?),
                b"numDim" => values = Some(parse_chartex_dimension(xml, &e, true)?),
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cx data")),
            _ => (),
        }
    }
    Ok((categories, values))
}

/// Parse a `cx:strDim` / `cx:numDim` element.
///
/// Unlike classic charts, chart-ex points hold their text directly in
/// `cx:pt`. Hierarchical dimensions (treemap/sunburst) keep all levels
/// in [`ChartDataSource::levels`] (innermost first, as in the
/// document), with the innermost level also mirrored into
/// [`ChartDataSource::values`].
fn parse_chartex_dimension<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    numeric: bool,
) -> Result<ChartDataSource, XlsxError> {
    let mut source = ChartDataSource::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"f" => source.formula = Some(read_text(xml, &e)?),
                b"lvl" => {
                    let level = parse_chartex_level(xml, &e, numeric)?;
                    if source.levels.is_empty() {
                        source.values = level.clone();
                    }
                    source.levels.push(level);
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cx dimension")),
            _ => (),
        }
    }
    Ok(source)
}

/// Parse one `cx:lvl` of a chart-ex dimension into a vector of points.
fn parse_chartex_level<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    numeric: bool,
) -> Result<Vec<Data>, XlsxError> {
    let pt_count: Option<usize> = parent.raw_attr(b"ptCount")?.and_then(parse_bytes);
    let mut points = Vec::new();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"pt" => {
                    let idx: Option<usize> = e.raw_attr(b"idx")?.and_then(parse_bytes);
                    let text = read_text(xml, &e)?;
                    let value = if numeric {
                        match text.trim().parse::<f64>() {
                            Ok(v) => Data::Float(v),
                            Err(_) => Data::String(text),
                        }
                    } else {
                        Data::String(text)
                    };
                    let idx = idx.unwrap_or(points.len());
                    if points.len() <= idx {
                        points.resize(idx + 1, Data::Empty);
                    }
                    points[idx] = value;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cx lvl")),
            _ => (),
        }
    }
    if let Some(count) = pt_count {
        if points.len() < count {
            points.resize(count, Data::Empty);
        }
    }
    Ok(points)
}

/// Parse a `cx:series` element into a single-series [`ChartGroup`].
fn parse_chartex_series<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    data: &std::collections::HashMap<String, (Option<ChartDataSource>, Option<ChartDataSource>)>,
    theme_colors: &[Color],
) -> Result<ChartGroup, XlsxError> {
    let mut group = ChartGroup::default();
    let mut series = ChartSeries::default();

    group.chart_type = match parent.raw_attr(b"layoutId")? {
        Some(b"boxWhisker") => ChartType::BoxWhisker,
        Some(b"clusteredColumn") => ChartType::Histogram,
        Some(b"funnel") => ChartType::Funnel,
        Some(b"paretoLine") => ChartType::Pareto,
        Some(b"regionMap") => ChartType::RegionMap,
        Some(b"sunburst") => ChartType::Sunburst,
        Some(b"treemap") => ChartType::Treemap,
        Some(b"waterfall") => ChartType::Waterfall,
        _ => ChartType::Unknown,
    };

    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"tx" | b"txData" => (),
                b"f" => {
                    series
                        .name
                        .get_or_insert_with(ChartDataSource::default)
                        .formula = Some(read_text(xml, &e)?);
                }
                b"v" => {
                    let text = read_text(xml, &e)?;
                    series
                        .name
                        .get_or_insert_with(ChartDataSource::default)
                        .values
                        .push(Data::String(text));
                }
                b"dataId" => {
                    if let Some(id) = e.raw_attr(b"val")? {
                        let id = decode_attr(&xml.decoder(), id)?;
                        if let Some((categories, values)) = data.get(&id) {
                            series.categories = categories.clone();
                            series.values = values.clone();
                        }
                    }
                    skip_element(xml, &e)?;
                }
                b"layoutPr" => {
                    series.chart_ex = Some(parse_chartex_layout(xml, &e)?);
                }
                b"spPr" => series.format = parse_shape_properties(xml, &e, theme_colors)?,
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cx series")),
            _ => (),
        }
    }

    group.series.push(series);
    Ok(group)
}

/// Parse a `cx:layoutPr` (series layout properties) element.
fn parse_chartex_layout<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
) -> Result<ChartExLayout, XlsxError> {
    let mut layout = ChartExLayout::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"parentLabelLayout" => {
                    layout.parent_label_layout = match e.raw_attr(b"val")? {
                        Some(b"none") => Some(ChartExParentLabelLayout::None),
                        Some(b"banner") => Some(ChartExParentLabelLayout::Banner),
                        Some(b"overlapping") => Some(ChartExParentLabelLayout::Overlapping),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"binning" => {
                    // "auto" thresholds stay None.
                    layout.overflow = e.raw_attr(b"overflow")?.and_then(parse_bytes);
                    layout.underflow = e.raw_attr(b"underflow")?.and_then(parse_bytes);
                    parse_chartex_binning(xml, &e, &mut layout)?;
                }
                b"statistics" => {
                    layout.quartile_method = match e.raw_attr(b"quartileMethod")? {
                        Some(b"inclusive") => Some(ChartExQuartileMethod::Inclusive),
                        Some(b"exclusive") => Some(ChartExQuartileMethod::Exclusive),
                        _ => None,
                    };
                    skip_element(xml, &e)?;
                }
                b"visibility" => {
                    let parse_flag = |v: Option<&[u8]>| {
                        v.map(|v| matches!(v, b"1" | b"true"))
                    };
                    layout.connector_lines = parse_flag(e.raw_attr(b"connectorLines")?);
                    layout.mean_line = parse_flag(e.raw_attr(b"meanLine")?);
                    layout.mean_marker = parse_flag(e.raw_attr(b"meanMarker")?);
                    layout.non_outliers = parse_flag(e.raw_attr(b"nonoutliers")?);
                    layout.outliers = parse_flag(e.raw_attr(b"outliers")?);
                    skip_element(xml, &e)?;
                }
                b"subtotals" => {
                    parse_chartex_subtotals(xml, &e, &mut layout.subtotals)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cx layoutPr")),
            _ => (),
        }
    }
    Ok(layout)
}

/// Parse the `cx:binSize` / `cx:binCount` children of `cx:binning`.
fn parse_chartex_binning<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    layout: &mut ChartExLayout,
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"binSize" => {
                    layout.bin_size = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                b"binCount" => {
                    layout.bin_count = val_attr(&e)?;
                    skip_element(xml, &e)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cx binning")),
            _ => (),
        }
    }
    Ok(())
}

/// Parse the `cx:idx` children of `cx:subtotals`.
fn parse_chartex_subtotals<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    subtotals: &mut Vec<u32>,
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"idx" => {
                    if let Some(idx) = val_attr(&e)? {
                        subtotals.push(idx);
                    }
                    skip_element(xml, &e)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cx subtotals")),
            _ => (),
        }
    }
    Ok(())
}

/// Parse a `cx:title` element.
fn parse_chartex_title<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartTitle, XlsxError> {
    let mut title = ChartTitle::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"tx" | b"txData" => (),
                b"rich" => {
                    let (rich, font, rotation) = parse_rich_text(xml, &e, theme_colors)?;
                    title.rich = Some(rich);
                    if title.font.is_none() {
                        title.font = font;
                    }
                    if title.text_rotation.is_none() {
                        title.text_rotation = rotation;
                    }
                }
                b"f" => title.formula = Some(read_text(xml, &e)?),
                b"v" => title.cached = Some(read_text(xml, &e)?),
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cx title")),
            _ => (),
        }
    }
    Ok(title)
}

/// Parse a `cx:axis` element.
///
/// Chart-ex axes are not explicitly typed; the type is inferred from the
/// scaling child (`cx:catScaling` or `cx:valScaling`).
fn parse_chartex_axis<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    theme_colors: &[Color],
) -> Result<ChartAxis, XlsxError> {
    let mut axis = ChartAxis::new(ChartAxisType::Category);
    axis.id = parent.raw_attr(b"id")?.and_then(parse_bytes);
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"catScaling" => {
                    axis.axis_type = ChartAxisType::Category;
                    axis.reverse = matches!(e.raw_attr(b"orientation")?, Some(b"maxMin"));
                    skip_element(xml, &e)?;
                }
                b"valScaling" => {
                    axis.axis_type = ChartAxisType::Value;
                    axis.reverse = matches!(e.raw_attr(b"orientation")?, Some(b"maxMin"));
                    axis.min = e.raw_attr(b"min")?.and_then(parse_bytes);
                    axis.max = e.raw_attr(b"max")?.and_then(parse_bytes);
                    skip_element(xml, &e)?;
                }
                b"title" => axis.title = Some(parse_chartex_title(xml, &e, theme_colors)?),
                b"numFmt" => {
                    if let Some(code) = e.raw_attr(b"formatCode")? {
                        axis.number_format = Some(decode_attr(&xml.decoder(), code)?);
                    }
                    skip_element(xml, &e)?;
                }
                b"majorGridlines" => {
                    axis.major_gridlines = true;
                    skip_element(xml, &e)?;
                }
                b"minorGridlines" => {
                    axis.minor_gridlines = true;
                    skip_element(xml, &e)?;
                }
                b"spPr" => axis.format = parse_shape_properties(xml, &e, theme_colors)?,
                b"txPr" => {
                    let (font, rotation) = parse_text_properties(xml, &e, theme_colors)?;
                    axis.font = font;
                    axis.text_rotation = rotation;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cx axis")),
            _ => (),
        }
    }
    Ok(axis)
}

// ---------------------------------------------------------------------------
// Drawing (xl/drawings/drawingN.xml) chart references
// ---------------------------------------------------------------------------

/// A chart reference found in a spreadsheet drawing: the relationship id of
/// the chart part plus the drawing object name and anchor.
pub(crate) struct DrawingChartRef {
    pub rel_id: String,
    pub name: Option<String>,
    pub position: ChartPosition,
}

/// Scan a drawing document for `graphicFrame` elements that embed charts.
pub(crate) fn drawing_chart_refs<RS: BufRead>(
    xml: &mut XmlReader<RS>,
) -> Result<Vec<DrawingChartRef>, XlsxError> {
    let mut refs = Vec::new();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"wsDr" => (),
                b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                    parse_drawing_anchor(xml, &e, &mut refs)?;
                }
                _ => skip_element(xml, &e)?,
            },
            Event::Eof => break,
            _ => (),
        }
    }
    Ok(refs)
}

/// Parse one drawing anchor, collecting any charts embedded in it.
fn parse_drawing_anchor<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
    refs: &mut Vec<DrawingChartRef>,
) -> Result<(), XlsxError> {
    let mut position = ChartPosition::default();
    let mut anchor_charts: Vec<(String, Option<String>)> = Vec::new();

    // Two-cell anchors default to editAs="twoCell" when the attribute is
    // absent; one-cell and absolute anchors have no editAs.
    if parent.local_name().as_ref() == b"twoCellAnchor" {
        position.edit_as = Some(match parent.raw_attr(b"editAs")? {
            Some(b"absolute") => ChartEditAs::Absolute,
            Some(b"oneCell") => ChartEditAs::OneCell,
            _ => ChartEditAs::TwoCell,
        });
    }

    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"from" => position.from = Some(parse_cell_anchor(xml, &e)?),
                b"to" => position.to = Some(parse_cell_anchor(xml, &e)?),
                // The absolute anchor position in EMUs.
                b"pos" => {
                    position.x = e.raw_attr(b"x")?.and_then(parse_bytes);
                    position.y = e.raw_attr(b"y")?.and_then(parse_bytes);
                    skip_element(xml, &e)?;
                }
                // The anchor-level extent; the graphicFrame's own a:ext is
                // consumed by parse_graphic_frame.
                b"ext" => {
                    position.width = e.raw_attr(b"cx")?.and_then(parse_bytes);
                    position.height = e.raw_attr(b"cy")?.and_then(parse_bytes);
                    skip_element(xml, &e)?;
                }
                b"graphicFrame" => {
                    if let Some(chart) = parse_graphic_frame(xml, &e)? {
                        anchor_charts.push(chart);
                    }
                }
                // Chart-ex frames are wrapped in mc:AlternateContent; the
                // mc:Fallback duplicates the frame as a picture, so only
                // descend into the mc:Choice branch.
                b"AlternateContent" | b"Choice" => (),
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("drawing anchor")),
            _ => (),
        }
    }

    for (rel_id, name) in anchor_charts {
        refs.push(DrawingChartRef {
            rel_id,
            name,
            position,
        });
    }
    Ok(())
}

/// Parse an `xdr:graphicFrame`, returning the chart relationship id and the
/// frame name if the frame embeds a chart.
fn parse_graphic_frame<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
) -> Result<Option<(String, Option<String>)>, XlsxError> {
    let mut name: Option<String> = None;
    let mut rel_id: Option<String> = None;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"cNvPr" => {
                    if let Some(n) = e.raw_attr(b"name")? {
                        name = Some(decode_attr(&xml.decoder(), n)?);
                    }
                    skip_element(xml, &e)?;
                }
                b"chart" => {
                    if let Some(id) = e.raw_attr_local(b"id")? {
                        rel_id = Some(decode_attr(&xml.decoder(), id)?);
                    }
                    skip_element(xml, &e)?;
                }
                // Skip the frame transform so its a:ext isn't mistaken for
                // the anchor extent.
                b"xfrm" => skip_element(xml, &e)?,
                b"nvGraphicFramePr" | b"graphic" | b"graphicData" => (),
                _ => skip_element(xml, &e)?,
            },
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("graphicFrame")),
            _ => (),
        }
    }
    Ok(rel_id.map(|id| (id, name)))
}

/// Parse an `xdr:from` / `xdr:to` cell anchor marker.
fn parse_cell_anchor<RS: BufRead>(
    xml: &mut XmlReader<RS>,
    parent: &BytesStart,
) -> Result<ChartCellAnchor, XlsxError> {
    let mut anchor = ChartCellAnchor::default();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let field = e.local_name().as_ref().to_vec();
                let text = read_text(xml, &e)?;
                let text = text.trim();
                match field.as_slice() {
                    b"col" => anchor.col = text.parse().unwrap_or(0),
                    b"row" => anchor.row = text.parse().unwrap_or(0),
                    b"colOff" => anchor.col_offset = text.parse().unwrap_or(0),
                    b"rowOff" => anchor.row_offset = text.parse().unwrap_or(0),
                    _ => (),
                }
            }
            Event::End(end) if end.name() == parent.name() => break,
            Event::Eof => return Err(XlsxError::XmlEof("cell anchor")),
            _ => (),
        }
    }
    Ok(anchor)
}

#[cfg(test)]
mod tests {
    use super::*;

    type BarCase = (&'static [u8], Option<&'static str>, Option<&'static str>, ChartType);

    fn resolve(
        tag: &[u8],
        bar_dir: Option<&str>,
        grouping: Option<&str>,
        scatter: Option<&str>,
        radar: Option<&str>,
        of_pie: Option<&str>,
        wireframe: bool,
    ) -> ChartType {
        resolve_chart_type(tag, bar_dir, grouping, scatter, radar, of_pie, wireframe)
    }

    #[test]
    fn resolve_bar_and_column_types() {
        let cases: &[BarCase] = &[
            (b"barChart", Some("col"), Some("clustered"), ChartType::Column),
            (b"barChart", Some("col"), Some("stacked"), ChartType::ColumnStacked),
            (b"barChart", Some("col"), Some("percentStacked"), ChartType::ColumnPercentStacked),
            (b"barChart", Some("bar"), Some("clustered"), ChartType::Bar),
            (b"barChart", Some("bar"), Some("stacked"), ChartType::BarStacked),
            (b"barChart", Some("bar"), Some("percentStacked"), ChartType::BarPercentStacked),
            (b"bar3DChart", Some("col"), Some("clustered"), ChartType::Column3D),
            (b"bar3DChart", Some("col"), Some("stacked"), ChartType::Column3DStacked),
            (b"bar3DChart", Some("col"), Some("percentStacked"), ChartType::Column3DPercentStacked),
            (b"bar3DChart", Some("col"), Some("standard"), ChartType::Column3DStandard),
            (b"bar3DChart", Some("bar"), Some("standard"), ChartType::Bar3DStandard),
            (b"bar3DChart", Some("bar"), Some("clustered"), ChartType::Bar3D),
            (b"bar3DChart", Some("bar"), Some("stacked"), ChartType::Bar3DStacked),
            (b"bar3DChart", Some("bar"), Some("percentStacked"), ChartType::Bar3DPercentStacked),
        ];
        for &(tag, dir, grouping, expected) in cases {
            assert_eq!(resolve(tag, dir, grouping, None, None, None, false), expected);
        }
    }

    #[test]
    fn resolve_line_area_pie_types() {
        let cases: &[(&[u8], Option<&str>, ChartType)] = &[
            (b"lineChart", Some("standard"), ChartType::Line),
            (b"lineChart", Some("stacked"), ChartType::LineStacked),
            (b"lineChart", Some("percentStacked"), ChartType::LinePercentStacked),
            (b"line3DChart", Some("standard"), ChartType::Line3D),
            (b"areaChart", Some("standard"), ChartType::Area),
            (b"areaChart", Some("stacked"), ChartType::AreaStacked),
            (b"areaChart", Some("percentStacked"), ChartType::AreaPercentStacked),
            (b"area3DChart", Some("standard"), ChartType::Area3D),
            (b"area3DChart", Some("stacked"), ChartType::Area3DStacked),
            (b"area3DChart", Some("percentStacked"), ChartType::Area3DPercentStacked),
            (b"pieChart", None, ChartType::Pie),
            (b"pie3DChart", None, ChartType::Pie3D),
            (b"doughnutChart", None, ChartType::Doughnut),
        ];
        for &(tag, grouping, expected) in cases {
            assert_eq!(resolve(tag, None, grouping, None, None, None, false), expected);
        }
    }

    #[test]
    fn resolve_scatter_radar_and_misc_types() {
        let scatter: &[(Option<&str>, ChartType)] = &[
            (Some("lineMarker"), ChartType::ScatterStraightWithMarkers),
            (Some("line"), ChartType::ScatterStraight),
            (Some("smoothMarker"), ChartType::ScatterSmoothWithMarkers),
            (Some("smooth"), ChartType::ScatterSmooth),
            (Some("marker"), ChartType::Scatter),
            (None, ChartType::Scatter),
        ];
        for &(style, expected) in scatter {
            assert_eq!(resolve(b"scatterChart", None, None, style, None, None, false), expected);
        }

        let radar: &[(Option<&str>, ChartType)] = &[
            (Some("standard"), ChartType::Radar),
            (Some("marker"), ChartType::RadarWithMarkers),
            (Some("filled"), ChartType::RadarFilled),
        ];
        for &(style, expected) in radar {
            assert_eq!(resolve(b"radarChart", None, None, None, style, None, false), expected);
        }

        assert_eq!(resolve(b"stockChart", None, None, None, None, None, false), ChartType::Stock);
        assert_eq!(resolve(b"bubbleChart", None, None, None, None, None, false), ChartType::Bubble);
        assert_eq!(
            resolve(b"ofPieChart", None, None, None, None, Some("pie"), false),
            ChartType::PieOfPie
        );
        assert_eq!(
            resolve(b"ofPieChart", None, None, None, None, Some("bar"), false),
            ChartType::BarOfPie
        );
        assert_eq!(resolve(b"surfaceChart", None, None, None, None, None, false), ChartType::Contour);
        assert_eq!(
            resolve(b"surfaceChart", None, None, None, None, None, true),
            ChartType::ContourWireframe
        );
        assert_eq!(
            resolve(b"surface3DChart", None, None, None, None, None, false),
            ChartType::Surface3D
        );
        assert_eq!(
            resolve(b"surface3DChart", None, None, None, None, None, true),
            ChartType::Surface3DWireframe
        );
        assert_eq!(resolve(b"unknownChart", None, None, None, None, None, false), ChartType::Unknown);
    }
}
