// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use quick_xml::{
    events::{attributes::Attribute, BytesStart, Event},
    name::QName,
    Reader,
};
use std::collections::BTreeMap;
use std::io::BufRead;

use crate::style::*;
use crate::XlsxError;

/// Parse color from XML attributes
fn parse_color(attributes: &[Attribute]) -> Result<Option<Color>, XlsxError> {
    for attr in attributes {
        match attr.key.as_ref() {
            b"rgb" => {
                let rgb_str = attr.value.as_ref();
                if rgb_str.len() == 6 {
                    let r = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[0..2]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid red color value"))?;
                    let g = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[2..4]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid green color value"))?;
                    let b = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[4..6]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid blue color value"))?;
                    return Ok(Some(Color::rgb(r, g, b)));
                }
            }
            b"theme" => {
                // TODO: Handle theme colors
                return Ok(None);
            }
            b"indexed" => {
                // TODO: Handle indexed colors
                return Ok(None);
            }
            _ => {}
        }
    }
    Ok(None)
}

/// Parse font weight from string
fn parse_font_weight(s: &str) -> FontWeight {
    match s {
        "bold" => FontWeight::Bold,
        _ => FontWeight::Normal,
    }
}

/// Parse font style from string
fn parse_font_style(s: &str) -> FontStyle {
    match s {
        "italic" => FontStyle::Italic,
        _ => FontStyle::Normal,
    }
}

/// Parse underline style from string
fn parse_underline_style(s: &str) -> UnderlineStyle {
    match s {
        "single" => UnderlineStyle::Single,
        "double" => UnderlineStyle::Double,
        "singleAccounting" => UnderlineStyle::SingleAccounting,
        "doubleAccounting" => UnderlineStyle::DoubleAccounting,
        _ => UnderlineStyle::None,
    }
}

/// Parse horizontal alignment from string
fn parse_horizontal_alignment(s: &str) -> HorizontalAlignment {
    match s {
        "left" => HorizontalAlignment::Left,
        "center" => HorizontalAlignment::Center,
        "right" => HorizontalAlignment::Right,
        "justify" => HorizontalAlignment::Justify,
        "distributed" => HorizontalAlignment::Distributed,
        "fill" => HorizontalAlignment::Fill,
        _ => HorizontalAlignment::General,
    }
}

/// Parse vertical alignment from string
fn parse_vertical_alignment(s: &str) -> VerticalAlignment {
    match s {
        "top" => VerticalAlignment::Top,
        "center" => VerticalAlignment::Center,
        "bottom" => VerticalAlignment::Bottom,
        "justify" => VerticalAlignment::Justify,
        "distributed" => VerticalAlignment::Distributed,
        _ => VerticalAlignment::Bottom,
    }
}

/// Parse fill pattern from string
fn parse_fill_pattern(s: &str) -> FillPattern {
    match s {
        "solid" => FillPattern::Solid,
        "darkGray" => FillPattern::DarkGray,
        "mediumGray" => FillPattern::MediumGray,
        "lightGray" => FillPattern::LightGray,
        "gray125" => FillPattern::Gray125,
        "gray0625" => FillPattern::Gray0625,
        "darkHorizontal" => FillPattern::DarkHorizontal,
        "darkVertical" => FillPattern::DarkVertical,
        "darkDown" => FillPattern::DarkDown,
        "darkUp" => FillPattern::DarkUp,
        "darkGrid" => FillPattern::DarkGrid,
        "darkTrellis" => FillPattern::DarkTrellis,
        "lightHorizontal" => FillPattern::LightHorizontal,
        "lightVertical" => FillPattern::LightVertical,
        "lightDown" => FillPattern::LightDown,
        "lightUp" => FillPattern::LightUp,
        "lightGrid" => FillPattern::LightGrid,
        "lightTrellis" => FillPattern::LightTrellis,
        _ => FillPattern::None,
    }
}

/// Parse border style from string
fn parse_border_style(s: &str) -> BorderStyle {
    match s {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "slantDashDot" => BorderStyle::SlantDashDot,
        _ => BorderStyle::None,
    }
}

/// Parse font element
fn parse_font<RS: BufRead>(
    xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Font, XlsxError> {
    let mut font = Font::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"name" => {
                    if let Some(name) = read_string(xml, QName(b"name"))? {
                        font = font.with_name(name);
                    }
                }
                b"sz" => {
                    if let Some(size_str) = read_string(xml, QName(b"sz"))? {
                        if let Ok(size) = size_str.parse::<f64>() {
                            font = font.with_size(size);
                        }
                    }
                }
                b"b" => {
                    font = font.with_weight(FontWeight::Bold);
                }
                b"i" => {
                    font = font.with_style(FontStyle::Italic);
                }
                b"u" => {
                    if let Some(underline_str) = read_string(xml, QName(b"u"))? {
                        font = font.with_underline(parse_underline_style(&underline_str));
                    }
                }
                b"strike" => {
                    font = font.with_strikethrough(true);
                }
                b"color" => {
                    if let Some(color) =
                        parse_color(&e.attributes().collect::<Result<Vec<_>, _>>()?)?
                    {
                        font = font.with_color(color);
                    }
                }
                b"family" => {
                    if let Some(family) = read_string(xml, QName(b"family"))? {
                        font = font.with_family(family);
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"font" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("font")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(font)
}

/// Parse fill element
fn parse_fill<RS: BufRead>(
    xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Fill, XlsxError> {
    let mut fill = Fill::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"patternFill" => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        match attr.key.as_ref() {
                            b"patternType" => {
                                let pattern_str = String::from_utf8_lossy(&attr.value);
                                fill = fill.with_pattern(parse_fill_pattern(&pattern_str));
                            }
                            _ => {}
                        }
                    }
                }
                b"fgColor" => {
                    if let Some(color) =
                        parse_color(&e.attributes().collect::<Result<Vec<_>, _>>()?)?
                    {
                        fill = fill.with_foreground_color(color);
                    }
                }
                b"bgColor" => {
                    if let Some(color) =
                        parse_color(&e.attributes().collect::<Result<Vec<_>, _>>()?)?
                    {
                        fill = fill.with_background_color(color);
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fill" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("fill")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(fill)
}

/// Parse border element
fn parse_border<RS: BufRead>(
    xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Borders, XlsxError> {
    let mut borders = Borders::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                    b"left" | b"right" | b"top" | b"bottom" | b"diagonal" => {
                        let mut style = BorderStyle::None;
                        let mut color = None;

                        for attr in e.attributes() {
                            let attr = attr?;
                            match attr.key.as_ref() {
                                b"style" => {
                                    let style_str = String::from_utf8_lossy(&attr.value);
                                    style = parse_border_style(&style_str);
                                }
                                _ => {}
                            }
                        }

                        if let Some(border_color) =
                            parse_color(&e.attributes().collect::<Result<Vec<_>, _>>()?)?
                        {
                            color = Some(border_color);
                        }

                        let border = if let Some(c) = color {
                            Border::with_color(style, c)
                        } else {
                            Border::new(style)
                        };

                        match e.local_name().as_ref() {
                            b"left" => borders.left = border,
                            b"right" => borders.right = border,
                            b"top" => borders.top = border,
                            b"bottom" => borders.bottom = border,
                            b"diagonal" => {
                                // Check if it's diagonal down or up
                                for attr in e.attributes() {
                                    let attr = attr?;
                                    if attr.key.as_ref() == b"diagonalDown" {
                                        borders.diagonal_down = border.clone();
                                    } else if attr.key.as_ref() == b"diagonalUp" {
                                        borders.diagonal_up = border.clone();
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"border" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("border")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(borders)
}

/// Parse alignment element
fn parse_alignment<RS: BufRead>(
    xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Alignment, XlsxError> {
    let mut alignment = Alignment::new();

    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"horizontal" => {
                let horizontal_str = String::from_utf8_lossy(&attr.value);
                alignment = alignment.with_horizontal(parse_horizontal_alignment(&horizontal_str));
            }
            b"vertical" => {
                let vertical_str = String::from_utf8_lossy(&attr.value);
                alignment = alignment.with_vertical(parse_vertical_alignment(&vertical_str));
            }
            b"wrapText" => {
                let wrap_str = String::from_utf8_lossy(&attr.value);
                if wrap_str == "1" || wrap_str == "true" {
                    alignment = alignment.with_wrap_text(true);
                }
            }
            b"textRotation" => {
                if let Ok(rotation) = String::from_utf8_lossy(&attr.value).parse::<u16>() {
                    alignment = alignment.with_text_rotation(TextRotation::Degrees(rotation));
                }
            }
            b"indent" => {
                if let Ok(indent) = String::from_utf8_lossy(&attr.value).parse::<u8>() {
                    alignment = alignment.with_indent(indent);
                }
            }
            b"shrinkToFit" => {
                let shrink_str = String::from_utf8_lossy(&attr.value);
                if shrink_str == "1" || shrink_str == "true" {
                    alignment = alignment.with_shrink_to_fit(true);
                }
            }
            _ => {}
        }
    }

    Ok(alignment)
}

/// Parse protection element
fn parse_protection<RS: BufRead>(
    xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Protection, XlsxError> {
    let mut protection = Protection::new();

    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"locked" => {
                let locked_str = String::from_utf8_lossy(&attr.value);
                if locked_str == "1" || locked_str == "true" {
                    protection = protection.with_locked(true);
                }
            }
            b"hidden" => {
                let hidden_str = String::from_utf8_lossy(&attr.value);
                if hidden_str == "1" || hidden_str == "true" {
                    protection = protection.with_hidden(true);
                }
            }
            _ => {}
        }
    }

    Ok(protection)
}

/// Parse a complete style (xf element)
pub fn parse_style<RS: BufRead>(
    xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Style, XlsxError> {
    let mut style = Style::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"font" => {
                    let font = parse_font(xml, e)?;
                    style = style.with_font(font);
                }
                b"fill" => {
                    let fill = parse_fill(xml, e)?;
                    style = style.with_fill(fill);
                }
                b"border" => {
                    let borders = parse_border(xml, e)?;
                    style = style.with_borders(borders);
                }
                b"alignment" => {
                    let alignment = parse_alignment(xml, e)?;
                    style = style.with_alignment(alignment);
                }
                b"protection" => {
                    let protection = parse_protection(xml, e)?;
                    style = style.with_protection(protection);
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"xf" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("xf")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(style)
}

/// Read string content from XML element
fn read_string<RS: BufRead>(
    xml: &mut Reader<RS>,
    closing: QName,
) -> Result<Option<String>, XlsxError> {
    let mut buf = Vec::new();
    let mut content = String::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Text(e)) => {
                content.push_str(&e.unescape()?);
            }
            Ok(Event::End(ref e)) if e.local_name() == closing.into() => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("string")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    if content.is_empty() {
        Ok(None)
    } else {
        Ok(Some(content))
    }
}
