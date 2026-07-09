// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

// Parsers for the style-related parts of an XLSX file: the theme color
// scheme in `xl/theme/theme1.xml` and the font, fill, border, alignment and
// protection elements of `xl/styles.xml`.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::io::BufRead;

use crate::attrs::{decode_attr, RawAttributes};
use crate::style::{
    Alignment, Border, BorderStyle, Borders, Color, Fill, FillPattern, Font, FontStyle, FontWeight,
    HorizontalAlignment, Protection, TextRotation, UnderlineStyle, VerticalAlignment,
};
use crate::xlsx::XlsxError;

// Default Office theme colors, indexed by Excel's `theme` attribute value.
//
// Excel swaps the first four entries relative to the `clrScheme` XML order:
// `theme="0"` -> lt1, `theme="1"` -> dk1, `theme="2"` -> lt2, `theme="3"` ->
// dk2.
pub(crate) fn default_theme_colors() -> Vec<Color> {
    vec![
        Color::rgb(255, 255, 255), // 0: lt1 (White).
        Color::rgb(0, 0, 0),       // 1: dk1 (Black).
        Color::rgb(238, 236, 225), // 2: lt2 (EEECE1).
        Color::rgb(31, 73, 125),   // 3: dk2 (1F497D).
        Color::rgb(79, 129, 189),  // 4: accent1.
        Color::rgb(192, 80, 77),   // 5: accent2.
        Color::rgb(155, 187, 89),  // 6: accent3.
        Color::rgb(128, 100, 162), // 7: accent4.
        Color::rgb(75, 172, 198),  // 8: accent5.
        Color::rgb(247, 150, 70),  // 9: accent6.
        Color::rgb(0, 0, 255),     // 10: hlink.
        Color::rgb(128, 0, 128),   // 11: folHlink.
    ]
}

// Resolve theme colors from the `clrScheme` element of `xl/theme/theme1.xml`.
//
// The `clrScheme` element lists colors in the order: dk1, lt1, dk2, lt2,
// accent1-accent6, hlink, folHlink. Excel swaps the first four when mapping
// to `theme` attribute indices, so: `theme="0"` -> lt1, `theme="1"` -> dk1,
// `theme="2"` -> lt2, `theme="3"` -> dk2.
//
// Returns the default Office theme colors if the color scheme is missing or
// incomplete. XML parse errors are propagated.
pub(crate) fn read_theme_colors<B: BufRead>(xml: &mut Reader<B>) -> Result<Vec<Color>, XlsxError> {
    const CLR_ELEMENTS: &[&str] = &[
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];

    let mut xml_order: Vec<Color> = Vec::new();
    let mut buf = Vec::new();
    let mut in_clr_scheme = false;

    // Set while inside one of the named clrScheme child elements, cleared
    // once a color has been recorded for it.
    let mut slot_open = false;

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == "clrScheme" {
                    in_clr_scheme = true;
                } else if in_clr_scheme {
                    if CLR_ELEMENTS.contains(&local.as_ref()) {
                        slot_open = true;
                    } else if slot_open && matches!(local.as_ref(), "srgbClr" | "sysClr") {
                        if let Some(color) = parse_theme_color(e)? {
                            xml_order.push(color);
                            slot_open = false;
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == "clrScheme" {
                    break;
                }
                if slot_open && CLR_ELEMENTS.contains(&local.as_ref()) {
                    // The slot closed without a resolvable color. Record a
                    // placeholder so later slots keep their correct indices.
                    xml_order.push(Color::rgb(0, 0, 0));
                    slot_open = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }

    if xml_order.len() >= 4 {
        // Apply Excel's swap of indices 0/1 and 2/3.
        xml_order.swap(0, 1);
        xml_order.swap(2, 3);
        Ok(xml_order)
    } else {
        Ok(default_theme_colors())
    }
}

// Parse an `srgbClr` or `sysClr` element into a color.
//
// For `srgbClr` the `val` attribute holds a 6-digit hex color. For `sysClr`
// the `lastClr` attribute (the resolved system color) is preferred over
// `val` (a system color name like "window").
fn parse_theme_color(e: &BytesStart) -> Result<Option<Color>, XlsxError> {
    let (val, last_clr) = get_attrs!(e, "val" => val, "lastClr" => last_clr)?;

    Ok(last_clr
        .and_then(parse_hex_color)
        .or_else(|| val.and_then(parse_hex_color)))
}

// Parse a 6-digit RGB or 8-digit ARGB hex color.
fn parse_hex_color(val: &str) -> Option<Color> {
    // `str::get()` returns `None` for a slice that is not on a char
    // boundary, so a non-ASCII value fails to parse rather than panicking.
    let hex_pair =
        |start: usize| -> Option<u8> { u8::from_str_radix(val.get(start..start + 2)?, 16).ok() };

    match val.len() {
        6 => Some(Color::rgb(hex_pair(0)?, hex_pair(2)?, hex_pair(4)?)),
        8 => Some(Color::new(
            hex_pair(0)?,
            hex_pair(2)?,
            hex_pair(4)?,
            hex_pair(6)?,
        )),
        _ => None,
    }
}

// Get a theme color, falling back to black for out-of-range indices.
fn get_theme_color(theme: u8, theme_colors: &[Color]) -> Color {
    theme_colors
        .get(theme as usize)
        .copied()
        .unwrap_or_else(|| Color::rgb(0, 0, 0))
}

// Get an indexed color from the default indexed color palette.
//
// The palette is defined in ECMA-376 section 18.8.27 (indexedColors).
// Indices 64 and 65 are the system foreground and background colors, which
// resolve to black and white respectively. Other out-of-palette indices
// resolve to black.
fn get_indexed_color(index: u8) -> Color {
    const PALETTE: [u32; 64] = [
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x000000,
        0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x800000, 0x008000,
        0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080, 0x9999FF, 0x993366, 0xFFFFCC,
        0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, 0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF,
        0x800080, 0x800000, 0x008080, 0x0000FF, 0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF,
        0xFF99CC, 0xCC99FF, 0xFFCC99, 0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600,
        0x666699, 0x969696, 0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399,
        0x333333,
    ];

    match index {
        0..=63 => Color::from_argb(0xFF00_0000 | PALETTE[index as usize]),
        // System foreground (auto).
        64 => Color::rgb(0, 0, 0),
        // System background.
        65 => Color::rgb(255, 255, 255),
        _ => Color::rgb(0, 0, 0),
    }
}

// Apply a tint to a color per ECMA-376 section 18.8.19.
//
// The tint is applied to the luminance of the color in HSL space, not to the
// RGB channels, so a tinted color keeps its hue. A negative tint darkens:
// `lum * (1 + tint)`. A positive tint lightens: `lum * (1 - tint) + tint`.
// The conversion is done in floating point and matches the colors in Excel's
// own theme color palette to within one unit per channel.
fn apply_tint(color: Color, tint: f64) -> Color {
    let (hue, sat, lum) = rgb_to_hsl(color.red, color.green, color.blue);

    let lum = if tint < 0.0 {
        lum * (1.0 + tint)
    } else {
        lum * (1.0 - tint) + tint
    };

    let (red, green, blue) = hsl_to_rgb(hue, sat, lum.clamp(0.0, 1.0));
    Color::new(color.alpha, red, green, blue)
}

// Convert an RGB color to hue, saturation and luminance values, each in the
// range 0.0 to 1.0.
fn rgb_to_hsl(red: u8, green: u8, blue: u8) -> (f64, f64, f64) {
    let r = f64::from(red) / 255.0;
    let g = f64::from(green) / 255.0;
    let b = f64::from(blue) / 255.0;

    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let lum = (max + min) / 2.0;

    if max == min {
        // Achromatic: the hue is undefined and the saturation is zero.
        return (0.0, 0.0, lum);
    }

    let delta = max - min;
    let sat = if lum > 0.5 {
        delta / (2.0 - max - min)
    } else {
        delta / (max + min)
    };

    let hue = if max == r {
        (g - b) / delta + if g < b { 6.0 } else { 0.0 }
    } else if max == g {
        (b - r) / delta + 2.0
    } else {
        (r - g) / delta + 4.0
    };

    (hue / 6.0, sat, lum)
}

// Convert hue, saturation and luminance values in the range 0.0 to 1.0 to
// RGB channels.
fn hsl_to_rgb(hue: f64, sat: f64, lum: f64) -> (u8, u8, u8) {
    let to_channel = |value: f64| (value * 255.0).round().clamp(0.0, 255.0) as u8;

    if sat == 0.0 {
        let grey = to_channel(lum);
        return (grey, grey, grey);
    }

    let q = if lum < 0.5 {
        lum * (1.0 + sat)
    } else {
        lum + sat - lum * sat
    };
    let p = 2.0 * lum - q;

    let hue_to_channel = |hue: f64| {
        let hue = hue.rem_euclid(1.0);
        let value = if hue < 1.0 / 6.0 {
            p + (q - p) * 6.0 * hue
        } else if hue < 0.5 {
            q
        } else if hue < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - hue) * 6.0
        } else {
            p
        };
        to_channel(value)
    };

    (
        hue_to_channel(hue + 1.0 / 3.0),
        hue_to_channel(hue),
        hue_to_channel(hue - 1.0 / 3.0),
    )
}

// Parse a color element (`color`, `fgColor`, `bgColor`, ...), resolving
// theme and indexed color references and applying the `tint` modifier when
// present.
pub(crate) fn parse_color(
    e: &BytesStart,
    theme_colors: &[Color],
) -> Result<Option<Color>, XlsxError> {
    let (rgb, theme, indexed, tint) = get_attrs!(
        e,
        "rgb" => rgb,
        "theme" => theme,
        "indexed" => indexed,
        "tint" => tint
    )?;

    let color = if let Some(rgb) = rgb {
        parse_hex_color(rgb)
    } else if let Some(theme) = theme {
        atoi_simd::parse::<u8, true, false>(theme.as_bytes())
            .ok()
            .map(|t| get_theme_color(t, theme_colors))
    } else if let Some(indexed) = indexed {
        atoi_simd::parse::<u8, true, false>(indexed.as_bytes())
            .ok()
            .map(get_indexed_color)
    } else {
        None
    };

    let tint = tint.and_then(|t| fast_float2::parse::<f64, _>(t.as_bytes()).ok());

    Ok(color.map(|c| match tint {
        Some(t) if t != 0.0 => apply_tint(c, t),
        _ => c,
    }))
}

// Parse an OOXML boolean attribute value.
fn parse_bool_attr(val: &str) -> bool {
    !matches!(val, "0" | "false")
}

// Parse a font weight value.
//
// Handles OOXML boolean values (`"1"`, `"true"`, `"0"`, `"false"`), named
// values (`"bold"`, `"normal"`), and numeric CSS-style weights.
fn parse_font_weight(val: &str) -> FontWeight {
    match val {
        "bold" | "700" | "1" | "true" => FontWeight::Bold,
        "normal" | "400" | "0" | "false" => FontWeight::Normal,
        _ => match atoi_simd::parse::<u16, true, false>(val.as_bytes()) {
            Ok(weight) if weight >= 600 => FontWeight::Bold,
            _ => FontWeight::Normal,
        },
    }
}

// Parse a font style value.
//
// Handles OOXML boolean values (`"1"`, `"true"`) in addition to named
// values (`"italic"`, `"oblique"`).
fn parse_font_style(val: &str) -> FontStyle {
    match val {
        "italic" | "oblique" | "1" | "true" => FontStyle::Italic,
        _ => FontStyle::Normal,
    }
}

fn parse_underline_style(val: &str) -> UnderlineStyle {
    match val {
        "single" => UnderlineStyle::Single,
        "double" => UnderlineStyle::Double,
        "singleAccounting" => UnderlineStyle::SingleAccounting,
        "doubleAccounting" => UnderlineStyle::DoubleAccounting,
        _ => UnderlineStyle::None,
    }
}

fn parse_horizontal_alignment(val: &str) -> HorizontalAlignment {
    match val {
        "left" => HorizontalAlignment::Left,
        "center" => HorizontalAlignment::Center,
        "centerContinuous" => HorizontalAlignment::CenterContinuous,
        "right" => HorizontalAlignment::Right,
        "justify" => HorizontalAlignment::Justify,
        "distributed" => HorizontalAlignment::Distributed,
        "fill" => HorizontalAlignment::Fill,
        _ => HorizontalAlignment::General,
    }
}

fn parse_vertical_alignment(val: &str) -> VerticalAlignment {
    match val {
        "top" => VerticalAlignment::Top,
        "center" => VerticalAlignment::Center,
        "bottom" => VerticalAlignment::Bottom,
        "justify" => VerticalAlignment::Justify,
        "distributed" => VerticalAlignment::Distributed,
        _ => VerticalAlignment::Bottom,
    }
}

fn parse_fill_pattern(val: &str) -> FillPattern {
    match val {
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

fn parse_border_style(val: &str) -> BorderStyle {
    match val {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "mediumDashDot" => BorderStyle::MediumDashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "mediumDashDotDot" => BorderStyle::MediumDashDotDot,
        "slantDashDot" => BorderStyle::SlantDashDot,
        _ => BorderStyle::None,
    }
}

// Parse a `<font>` element from `xl/styles.xml`.
//
// The reader must be positioned just after the `<font>` start element; this
// function consumes events up to and including the matching end element.
pub(crate) fn parse_font<B: BufRead>(
    xml: &mut Reader<B>,
    theme_colors: &[Color],
) -> Result<Font, XlsxError> {
    let mut font = Font::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                "name" => {
                    if let Some(val) = e.raw_attr("val")? {
                        font = font.set_name(decode_attr(val)?);
                    }
                }
                "sz" => {
                    if let Some(val) = e.raw_attr("val")? {
                        if let Ok(size) = fast_float2::parse::<f64, _>(val.as_bytes()) {
                            font = font.set_size(size);
                        }
                    }
                }
                "b" => {
                    let weight = match e.raw_attr("val")? {
                        Some(val) => parse_font_weight(val),
                        None => FontWeight::Bold,
                    };
                    font = font.set_weight(weight);
                }
                "i" => {
                    let style = match e.raw_attr("val")? {
                        Some(val) => parse_font_style(val),
                        None => FontStyle::Italic,
                    };
                    font = font.set_style(style);
                }
                "u" => {
                    let underline = match e.raw_attr("val")? {
                        Some(val) => parse_underline_style(val),
                        None => UnderlineStyle::Single,
                    };
                    font = font.set_underline(underline);
                }
                "strike" => {
                    let strike = match e.raw_attr("val")? {
                        Some(val) => parse_bool_attr(val),
                        None => true,
                    };
                    font = font.set_strikethrough(strike);
                }
                "color" => {
                    if let Some(color) = parse_color(e, theme_colors)? {
                        font = font.set_color(color);
                    }
                }
                "family" => {
                    if let Some(val) = e.raw_attr("val")? {
                        if let Ok(family) = atoi_simd::parse::<u8, true, false>(val.as_bytes()) {
                            font = font.set_family(family);
                        }
                    }
                }
                _ => (),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == "font" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("font")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }

    Ok(font)
}

// Parse a `<fill>` element from `xl/styles.xml`.
//
// The reader must be positioned just after the `<fill>` start element; this
// function consumes events up to and including the matching end element.
// Gradient fills are not supported and result in a default (empty) fill.
pub(crate) fn parse_fill<B: BufRead>(
    xml: &mut Reader<B>,
    theme_colors: &[Color],
) -> Result<Fill, XlsxError> {
    let mut fill = Fill::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                "patternFill" => {
                    if let Some(val) = e.raw_attr("patternType")? {
                        fill = fill.set_pattern(parse_fill_pattern(val));
                    }
                }
                "fgColor" => {
                    if let Some(color) = parse_color(e, theme_colors)? {
                        fill = fill.set_foreground_color(color);
                    }
                }
                "bgColor" => {
                    if let Some(color) = parse_color(e, theme_colors)? {
                        fill = fill.set_background_color(color);
                    }
                }
                _ => (),
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == "fill" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("fill")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }

    Ok(fill)
}

// Parse a `<border>` element from `xl/styles.xml`.
//
// The reader must be positioned just after the `<border>` start element;
// this function consumes events up to and including the matching end
// element. The `diagonalUp`/`diagonalDown` attributes of the start element
// select which diagonal borders the `<diagonal>` child applies to.
pub(crate) fn parse_border<B: BufRead>(
    xml: &mut Reader<B>,
    start_elem: &BytesStart,
    theme_colors: &[Color],
) -> Result<Borders, XlsxError> {
    let mut borders = Borders::new();
    let mut buf = Vec::new();
    let mut side_buf = Vec::new();

    let (diagonal_down, diagonal_up) = get_attrs!(
        start_elem,
        "diagonalDown" => diagonal_down,
        "diagonalUp" => diagonal_up
    )?;
    let has_diagonal_down = diagonal_down.map(parse_bool_attr).unwrap_or(false);
    let has_diagonal_up = diagonal_up.map(parse_bool_attr).unwrap_or(false);

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(ref event @ (Event::Start(ref e) | Event::Empty(ref e))) => {
                let side = e.local_name();
                let side = side.as_ref();
                if !matches!(side, "left" | "right" | "top" | "bottom" | "diagonal") {
                    continue;
                }

                let style = match e.raw_attr("style")? {
                    Some(val) => parse_border_style(val),
                    None => BorderStyle::None,
                };

                // A non-empty side element can contain a <color> child.
                let mut color = None;
                if matches!(event, Event::Start(_)) {
                    loop {
                        side_buf.clear();
                        match xml.read_event_into(&mut side_buf) {
                            Ok(Event::Start(ref c) | Event::Empty(ref c))
                                if c.local_name().as_ref() == "color" =>
                            {
                                if let Some(border_color) = parse_color(c, theme_colors)? {
                                    color = Some(border_color);
                                }
                            }
                            Ok(Event::End(ref end)) if end.local_name().as_ref() == side => {
                                break;
                            }
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("border side")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                }

                let border = match color {
                    Some(color) => Border::with_color(style, color),
                    None => Border::new(style),
                };

                match side {
                    "left" => borders.left = border,
                    "right" => borders.right = border,
                    "top" => borders.top = border,
                    "bottom" => borders.bottom = border,
                    "diagonal" => {
                        if has_diagonal_down {
                            borders.diagonal_down = border.clone();
                        }
                        if has_diagonal_up {
                            borders.diagonal_up = border;
                        }
                    }
                    _ => (),
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == "border" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("border")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }

    Ok(borders)
}

// Parse an `<alignment>` element's attributes.
pub(crate) fn parse_alignment(e: &BytesStart) -> Result<Alignment, XlsxError> {
    let mut alignment = Alignment::new();

    let (horizontal, vertical, wrap_text, text_rotation, indent, shrink_to_fit) = get_attrs!(
        e,
        "horizontal" => horizontal,
        "vertical" => vertical,
        "wrapText" => wrap_text,
        "textRotation" => text_rotation,
        "indent" => indent,
        "shrinkToFit" => shrink_to_fit
    )?;

    if let Some(val) = horizontal {
        alignment = alignment.set_horizontal(parse_horizontal_alignment(val));
    }
    if let Some(val) = vertical {
        alignment = alignment.set_vertical(parse_vertical_alignment(val));
    }
    if let Some(val) = wrap_text {
        alignment = alignment.set_wrap_text(parse_bool_attr(val));
    }
    if let Some(val) = text_rotation {
        if let Ok(rotation) = atoi_simd::parse::<u16, true, false>(val.as_bytes()) {
            let rotation = if rotation == 255 {
                TextRotation::Stacked
            } else {
                TextRotation::Degrees(rotation)
            };
            alignment = alignment.set_text_rotation(rotation);
        }
    }
    if let Some(val) = indent {
        if let Ok(indent) = atoi_simd::parse::<u8, true, false>(val.as_bytes()) {
            alignment = alignment.set_indent(indent);
        }
    }
    if let Some(val) = shrink_to_fit {
        alignment = alignment.set_shrink_to_fit(parse_bool_attr(val));
    }

    Ok(alignment)
}

// Parse a `<protection>` element's attributes.
//
// Cells are locked by default per the OOXML spec.
pub(crate) fn parse_protection(e: &BytesStart) -> Result<Protection, XlsxError> {
    let (locked, hidden) = get_attrs!(e, "locked" => locked, "hidden" => hidden)?;

    Ok(Protection::new()
        .set_locked(locked.map(parse_bool_attr).unwrap_or(true))
        .set_hidden(hidden.map(parse_bool_attr).unwrap_or(false)))
}

#[cfg(test)]
mod tests {
    use super::*;

    // Build a BufRead-backed reader over an XML snippet and consume the
    // first (start) event so the parse_* functions see only the content.
    fn reader_inside(xml_text: &'static str) -> Reader<&'static [u8]> {
        let mut reader = Reader::from_reader(xml_text.as_bytes());
        let mut buf = Vec::new();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(_)) => (),
            other => panic!("expected start event, got {other:?}"),
        }
        reader
    }

    #[test]
    fn test_parse_font_weight_values() {
        assert_eq!(parse_font_weight("1"), FontWeight::Bold);
        assert_eq!(parse_font_weight("true"), FontWeight::Bold);
        assert_eq!(parse_font_weight("bold"), FontWeight::Bold);
        assert_eq!(parse_font_weight("700"), FontWeight::Bold);
        assert_eq!(parse_font_weight("600"), FontWeight::Bold);
        assert_eq!(parse_font_weight("0"), FontWeight::Normal);
        assert_eq!(parse_font_weight("false"), FontWeight::Normal);
        assert_eq!(parse_font_weight("normal"), FontWeight::Normal);
        assert_eq!(parse_font_weight("400"), FontWeight::Normal);
        assert_eq!(parse_font_weight("599"), FontWeight::Normal);
        assert_eq!(parse_font_weight("junk"), FontWeight::Normal);
    }

    #[test]
    fn test_parse_font_style_values() {
        assert_eq!(parse_font_style("italic"), FontStyle::Italic);
        assert_eq!(parse_font_style("oblique"), FontStyle::Italic);
        assert_eq!(parse_font_style("1"), FontStyle::Italic);
        assert_eq!(parse_font_style("true"), FontStyle::Italic);
        assert_eq!(parse_font_style("0"), FontStyle::Normal);
        assert_eq!(parse_font_style("false"), FontStyle::Normal);
        assert_eq!(parse_font_style("normal"), FontStyle::Normal);
    }

    #[test]
    fn test_indexed_color_palette() {
        // ECMA-376 18.8.27 palette entries (0-based, unlike VBA ColorIndex).
        assert_eq!(get_indexed_color(0), Color::rgb(0, 0, 0));
        assert_eq!(get_indexed_color(1), Color::rgb(255, 255, 255));
        assert_eq!(get_indexed_color(2), Color::rgb(255, 0, 0));
        assert_eq!(get_indexed_color(3), Color::rgb(0, 255, 0));
        assert_eq!(get_indexed_color(22), Color::rgb(192, 192, 192));
        assert_eq!(get_indexed_color(63), Color::rgb(51, 51, 51));
        // System foreground/background.
        assert_eq!(get_indexed_color(64), Color::rgb(0, 0, 0));
        assert_eq!(get_indexed_color(65), Color::rgb(255, 255, 255));
        // Out of palette.
        assert_eq!(get_indexed_color(99), Color::rgb(0, 0, 0));
    }

    #[test]
    fn test_parse_hex_color() {
        assert_eq!(parse_hex_color("FF8040"), Some(Color::rgb(255, 128, 64)));
        assert_eq!(
            parse_hex_color("80FF8040"),
            Some(Color::new(128, 255, 128, 64))
        );
        assert_eq!(parse_hex_color("FFF"), None);
        assert_eq!(parse_hex_color("GGGGGG"), None);
    }

    #[test]
    fn test_apply_tint() {
        // Greys are unaffected by the HSL conversion so the results are
        // exact: a positive tint lightens and a negative tint darkens.
        assert_eq!(
            apply_tint(Color::rgb(0, 0, 0), 0.5),
            Color::rgb(128, 128, 128)
        );
        assert_eq!(
            apply_tint(Color::rgb(255, 255, 255), -0.5),
            Color::rgb(128, 128, 128)
        );
        assert_eq!(
            apply_tint(Color::rgb(255, 255, 255), -0.25),
            Color::rgb(191, 191, 191)
        );

        // Alpha is preserved.
        let tinted = apply_tint(Color::new(10, 100, 100, 100), 0.0);
        assert_eq!(tinted.alpha, 10);
    }

    // Check a tinted color against the value shown in Excel's theme color
    // palette, allowing a difference of one per channel for rounding.
    fn assert_tint_matches_excel(base: Color, tint: f64, expected: Color) {
        let tinted = apply_tint(base, tint);
        for (got, want) in [
            (tinted.red, expected.red),
            (tinted.green, expected.green),
            (tinted.blue, expected.blue),
        ] {
            assert!(
                got.abs_diff(want) <= 1,
                "tint {tint} of {base} gave {tinted}, expected {expected}"
            );
        }
    }

    #[test]
    fn test_apply_tint_matches_excel_palette() {
        // The tints are the values Excel writes for its "Lighter"/"Darker"
        // palette entries and the expected colors are the palette swatches
        // of the Office 2007 theme. The tint applies to the HSL luminance so
        // a tinted color keeps its hue: applying it per RGB channel would
        // turn the tan lt2 color into a neutral grey.
        let lt2 = Color::rgb(0xEE, 0xEC, 0xE1);
        assert_tint_matches_excel(lt2, -0.0999786370433668, Color::rgb(0xDD, 0xD9, 0xC4));
        assert_tint_matches_excel(lt2, -0.249977111117893, Color::rgb(0xC4, 0xBD, 0x97));
        assert_tint_matches_excel(lt2, -0.499984740745262, Color::rgb(0x93, 0x89, 0x53));

        let dk2 = Color::rgb(0x1F, 0x49, 0x7D);
        assert_tint_matches_excel(dk2, 0.799981688894314, Color::rgb(0xC6, 0xD9, 0xF1));
        assert_tint_matches_excel(dk2, 0.599993896298105, Color::rgb(0x8D, 0xB4, 0xE2));
        assert_tint_matches_excel(dk2, 0.399975585192419, Color::rgb(0x54, 0x8D, 0xD4));
        assert_tint_matches_excel(dk2, -0.249977111117893, Color::rgb(0x17, 0x37, 0x5E));

        let accent1 = Color::rgb(0x4F, 0x81, 0xBD);
        assert_tint_matches_excel(accent1, 0.799981688894314, Color::rgb(0xDC, 0xE6, 0xF1));
        assert_tint_matches_excel(accent1, 0.399975585192419, Color::rgb(0x95, 0xB3, 0xD7));
        assert_tint_matches_excel(accent1, -0.249977111117893, Color::rgb(0x36, 0x60, 0x92));
        assert_tint_matches_excel(accent1, -0.499984740745262, Color::rgb(0x24, 0x40, 0x61));

        let accent6 = Color::rgb(0xF7, 0x96, 0x46);
        assert_tint_matches_excel(accent6, -0.499984740745262, Color::rgb(0x97, 0x48, 0x06));

        let white = Color::rgb(255, 255, 255);
        assert_tint_matches_excel(white, -0.0499893185216834, Color::rgb(0xF2, 0xF2, 0xF2));
        assert_tint_matches_excel(white, -0.499984740745262, Color::rgb(0x80, 0x80, 0x80));

        let black = Color::rgb(0, 0, 0);
        assert_tint_matches_excel(black, 0.249977111117893, Color::rgb(0x40, 0x40, 0x40));
        assert_tint_matches_excel(black, 0.499984740745262, Color::rgb(0x7F, 0x7F, 0x7F));
    }

    #[test]
    fn test_parse_border_style_values() {
        assert_eq!(parse_border_style("thin"), BorderStyle::Thin);
        assert_eq!(
            parse_border_style("mediumDashDot"),
            BorderStyle::MediumDashDot
        );
        assert_eq!(
            parse_border_style("mediumDashDotDot"),
            BorderStyle::MediumDashDotDot
        );
        assert_eq!(
            parse_border_style("slantDashDot"),
            BorderStyle::SlantDashDot
        );
        assert_eq!(parse_border_style("none"), BorderStyle::None);
        assert_eq!(parse_border_style("junk"), BorderStyle::None);
    }

    #[test]
    fn test_parse_horizontal_alignment_values() {
        assert_eq!(
            parse_horizontal_alignment("center"),
            HorizontalAlignment::Center
        );
        assert_eq!(
            parse_horizontal_alignment("centerContinuous"),
            HorizontalAlignment::CenterContinuous
        );
        assert_eq!(
            parse_horizontal_alignment("general"),
            HorizontalAlignment::General
        );
        assert_eq!(
            parse_horizontal_alignment("junk"),
            HorizontalAlignment::General
        );
    }

    #[test]
    fn test_parse_font_element() {
        let mut xml = reader_inside(
            r#"<font><b/><i val="1"/><u/><strike/><sz val="12.5"/><name val="Arial"/><family val="2"/><color rgb="FF00FF00"/></font>"#,
        );
        let font = parse_font(&mut xml, &default_theme_colors()).unwrap();

        assert!(font.is_bold());
        assert!(font.is_italic());
        assert_eq!(font.underline, UnderlineStyle::Single);
        assert!(font.has_strikethrough());
        assert_eq!(font.size, Some(12.5));
        assert_eq!(font.name.as_deref(), Some("Arial"));
        assert_eq!(font.family, Some(2));
        assert_eq!(font.color, Some(Color::rgb(0, 255, 0)));
    }

    #[test]
    fn test_parse_font_explicit_off_values() {
        let mut xml = reader_inside(r#"<font><b val="0"/><u val="none"/></font>"#);
        let font = parse_font(&mut xml, &default_theme_colors()).unwrap();

        assert!(!font.is_bold());
        assert_eq!(font.underline, UnderlineStyle::None);
    }

    #[test]
    fn test_parse_fill_element() {
        let mut xml = reader_inside(
            r#"<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>"#,
        );
        let fill = parse_fill(&mut xml, &default_theme_colors()).unwrap();

        assert_eq!(fill.pattern, FillPattern::Solid);
        assert_eq!(fill.foreground_color, Some(Color::rgb(255, 255, 0)));
        assert_eq!(fill.background_color, Some(Color::rgb(0, 0, 0)));
        assert!(fill.is_visible());
    }

    #[test]
    fn test_parse_border_element() {
        let mut xml = Reader::from_reader(
            r#"<border diagonalUp="1"><left style="thin"><color rgb="FFFF0000"/></left><right/><top style="thick"/><bottom/><diagonal style="dashed"/></border>"#
                .as_bytes(),
        );
        let mut buf = Vec::new();
        let start = match xml.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => e.into_owned(),
            other => panic!("expected start event, got {other:?}"),
        };
        let borders = parse_border(&mut xml, &start, &default_theme_colors()).unwrap();

        assert_eq!(borders.left.style, BorderStyle::Thin);
        assert_eq!(borders.left.color, Some(Color::rgb(255, 0, 0)));
        assert_eq!(borders.right.style, BorderStyle::None);
        assert_eq!(borders.top.style, BorderStyle::Thick);
        assert_eq!(borders.diagonal_up.style, BorderStyle::Dashed);
        assert_eq!(borders.diagonal_down.style, BorderStyle::None);
    }

    #[test]
    fn test_parse_alignment_element() {
        let xml_text = r#"<alignment horizontal="center" vertical="top" wrapText="1" textRotation="255" indent="2" shrinkToFit="true"/>"#;
        let mut xml = Reader::from_reader(xml_text.as_bytes());
        let mut buf = Vec::new();
        let e = match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => e.into_owned(),
            other => panic!("expected empty event, got {other:?}"),
        };
        let alignment = parse_alignment(&e).unwrap();

        assert_eq!(alignment.horizontal, HorizontalAlignment::Center);
        assert_eq!(alignment.vertical, VerticalAlignment::Top);
        assert!(alignment.wrap_text);
        assert_eq!(alignment.text_rotation, TextRotation::Stacked);
        assert_eq!(alignment.indent, Some(2));
        assert!(alignment.shrink_to_fit);
    }

    #[test]
    fn test_parse_protection_defaults() {
        let xml_text = r#"<protection/>"#;
        let mut xml = Reader::from_reader(xml_text.as_bytes());
        let mut buf = Vec::new();
        let e = match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => e.into_owned(),
            other => panic!("expected empty event, got {other:?}"),
        };
        let protection = parse_protection(&e).unwrap();

        // Locked by default per the OOXML spec.
        assert!(protection.locked);
        assert!(!protection.hidden);
    }

    #[test]
    fn test_read_theme_colors_swaps_first_four() {
        let xml_text = r#"<a:clrScheme xmlns:a="x" name="Office">
            <a:dk1><a:sysClr val="windowText" lastClr="000001"/></a:dk1>
            <a:lt1><a:sysClr val="window" lastClr="FFFFFE"/></a:lt1>
            <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
            <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
            <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
        </a:clrScheme>"#;
        let mut xml = Reader::from_reader(xml_text.as_bytes());
        let colors = read_theme_colors(&mut xml).unwrap();

        // Swapped: 0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2.
        assert_eq!(colors[0], Color::rgb(255, 255, 254));
        assert_eq!(colors[1], Color::rgb(0, 0, 1));
        assert_eq!(colors[2], Color::rgb(238, 236, 225));
        assert_eq!(colors[3], Color::rgb(31, 73, 125));
        assert_eq!(colors[4], Color::rgb(79, 129, 189));
    }

    #[test]
    fn test_read_theme_colors_incomplete_falls_back() {
        let xml_text =
            r#"<a:clrScheme xmlns:a="x"><a:dk1><a:srgbClr val="000000"/></a:dk1></a:clrScheme>"#;
        let mut xml = Reader::from_reader(xml_text.as_bytes());
        let colors = read_theme_colors(&mut xml).unwrap();

        assert_eq!(colors, default_theme_colors());
    }

    #[test]
    fn test_read_theme_colors_malformed_xml_is_an_error() {
        // XML errors are propagated to the caller, which falls back to the
        // default theme colors.
        let xml_text = r#"<a:clrScheme xmlns:a="x"><a:dk1><a:srgbClr val="000000"/></a:dk1><"#;
        let mut xml = Reader::from_reader(xml_text.as_bytes());

        assert!(read_theme_colors(&mut xml).is_err());
    }

    #[test]
    fn test_sys_clr_prefers_last_clr() {
        let xml_text = r#"<sysClr val="window" lastClr="ABCDEF"/>"#;
        let mut xml = Reader::from_reader(xml_text.as_bytes());
        let mut buf = Vec::new();
        let e = match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => e.into_owned(),
            other => panic!("expected empty event, got {other:?}"),
        };

        assert_eq!(
            parse_theme_color(&e).unwrap(),
            Some(Color::rgb(0xAB, 0xCD, 0xEF))
        );
    }
}
