//! xlsx 快速路径：zip 流式解压 → 按 `<row` 边界切块 → rayon 并行解析 → 稀疏 [`Cell`] 列表。
//!
//! 只处理无命名空间前缀的标准 SpreadsheetML。解析器是严格的：任何不认识的结构
//! （带前缀的标签、`<v>` 里的注释/CDATA、单引号属性等）都返回 [`FastErr`]，
//! 调用方随即回退到 quick-xml 路径（`Xlsx::worksheet_range_ref`），保证行为等价。
//! 单元格值的类型转换复用现有的 `strings` / `formats` / [`format_excel_f64`]，
//! 文本语义（EOL 规范化、实体解码、`xml:space` trim、`_x000D_` 转义、富文本拼接）
//! 与 `read_string_with_bufs` / `read_v` 一一对应。

use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};
use std::sync::mpsc;

use memchr::{memchr, memmem};
use rayon::prelude::*;
use zip::ZipArchive;

use crate::datatype::Data;
use crate::formats::{format_excel_f64, CellFormat};
use crate::utils::{cached_zip_path, unescape_entity_to_buffer, unescape_xml};
use crate::{Cell, CellErrorType, HeaderRow, Range};

/// 任何 fast path 处理不了的情况：整体回退旧路径。
#[derive(Debug)]
struct FastErr;

type Res<T> = Result<T, FastErr>;

const BLOCK_SIZE: usize = 4 * 1024 * 1024;
/// 一直找不到 `<row` 切点时 carry 的上限，超过就放弃（防止带前缀的文档把内存吃爆）。
const MAX_CARRY: usize = 64 * 1024 * 1024;

// ---------------------------------------------------------------- 块读取

/// 按完整 `<row` 起点切块读取解压后的 sheet XML；首块带上 probe 读过的头部字节。
struct BlockReader<R: Read> {
    rd: R,
    carry: Vec<u8>,
    eof: bool,
}

impl<R: Read> BlockReader<R> {
    fn new(rd: R, head: Vec<u8>) -> Self {
        BlockReader {
            rd,
            carry: head,
            eof: false,
        }
    }

    /// 取下一块（≥ `target` 字节，尾部对齐到 `<row` 起点）；读完返回 None。
    /// `Err(())` = 长期没有切点（多半是带前缀的文档），应回退。
    fn next_block(&mut self, target: usize) -> Res<Option<Vec<u8>>> {
        loop {
            if self.eof {
                return Ok(if self.carry.is_empty() {
                    None
                } else {
                    Some(std::mem::take(&mut self.carry))
                });
            }
            if self.carry.len() >= target {
                if let Some(cut) = self.cut_point() {
                    let tail = self.carry.split_off(cut);
                    return Ok(Some(std::mem::replace(&mut self.carry, tail)));
                }
                if self.carry.len() >= MAX_CARRY {
                    return Err(FastErr);
                }
            }
            let old = self.carry.len();
            let want = (target / 4).max(64 * 1024);
            self.carry.resize(old + want, 0);
            let mut got = 0;
            while got < want {
                match self.rd.read(&mut self.carry[old + got..]) {
                    Ok(0) | Err(_) => {
                        self.eof = true;
                        break;
                    }
                    Ok(k) => got += k,
                }
            }
            self.carry.truncate(old + got);
        }
    }

    /// carry 里最后一个合法的 `<row` 起点切点（后面必须跟空白或 `>`、`/`）。
    fn cut_point(&self) -> Option<usize> {
        let mut end = self.carry.len();
        loop {
            let cut = memmem::rfind(&self.carry[..end], b"<row")?;
            let next = self.carry.get(cut + 4);
            if matches!(
                next,
                Some(b' ') | Some(b'\t') | Some(b'\r') | Some(b'\n') | Some(b'>') | Some(b'/')
            ) {
                // 切点不能是 carry 开头，否则原地不动死循环
                if cut > 0 {
                    return Some(cut);
                }
            }
            end = cut;
        }
    }
}

// ---------------------------------------------------------------- 标签与属性

struct Tag<'a> {
    name: &'a [u8],
    attrs: &'a [u8],
    self_closing: bool,
}

/// 解析 start tag（`rest` 以 `<` 开头且不是 `</`），返回 tag 和消耗的字节数。
fn parse_start_tag(rest: &[u8]) -> Res<(Tag<'_>, usize)> {
    let gt = memchr(b'>', rest).ok_or(FastErr)?;
    let inner = &rest[1..gt];
    let (self_closing, inner) = if inner.last() == Some(&b'/') {
        (true, &inner[..inner.len() - 1])
    } else {
        (false, inner)
    };
    let name_end = inner
        .iter()
        .position(|b| b.is_ascii_whitespace())
        .unwrap_or(inner.len());
    Ok((
        Tag {
            name: &inner[..name_end],
            attrs: &inner[name_end..],
            self_closing,
        },
        gt + 1,
    ))
}

/// 解析 end tag（`rest` 以 `</` 开头），返回标签名和消耗的字节数。
fn parse_end_tag(rest: &[u8]) -> Res<(&[u8], usize)> {
    let gt = memchr(b'>', rest).ok_or(FastErr)?;
    Ok((&rest[2..gt], gt + 1))
}

/// 在 attrs 里找 ` name="value"`（XML 属性名前必须是空白）。
fn attr<'a>(attrs: &'a [u8], name: &[u8]) -> Option<&'a [u8]> {
    let mut from = 0;
    loop {
        let i = memmem::find(&attrs[from..], name)?;
        let at = from + i;
        let ok_before = at > 0 && attrs[at - 1].is_ascii_whitespace();
        let rest = &attrs[at + name.len()..];
        if ok_before && rest.starts_with(b"=\"") {
            let end = memchr(b'"', &rest[2..])?;
            return Some(&rest[2..2 + end]);
        }
        from = at + 1;
    }
}

// ---------------------------------------------------------------- 文本处理

/// `\r\n` / `\r` 规范成 `\n`（XML 1.0 规则，对齐 quick-xml 的 `xml10_content()`）。
fn push_normalized(s: &str, out: &mut String) {
    if !s.contains('\r') {
        out.push_str(s);
        return;
    }
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '\r' {
            if chars.peek() == Some(&'\n') {
                chars.next();
            }
            out.push('\n');
        } else {
            out.push(c);
        }
    }
}

/// 一段纯文本：`&…;` 实体解码后推进 `out`；无法识别的实体报错（legacy 会报错，
/// 这里回退交给它）。
fn push_text_with_refs(s: &str, out: &mut String) -> Res<()> {
    let mut rest = s;
    loop {
        match rest.find('&') {
            None => {
                push_normalized(rest, out);
                return Ok(());
            }
            Some(i) => {
                push_normalized(&rest[..i], out);
                let after = &rest[i + 1..];
                let semi = after.find(';').filter(|&e| e <= 10).ok_or(FastErr)?;
                let ent = &after[..semi];
                unescape_entity_to_buffer(&quick_xml::events::BytesRef::new(ent), out)
                    .map_err(|_| FastErr)?;
                rest = &after[semi + 1..];
            }
        }
    }
}

/// 把已截取的元素内容当作 legacy 的文本事件序列推进 `out`：文本段做 EOL 规范化、
/// `&…;` 实体解码；注释跳过；CDATA 按 `push_cdata` 决定推进（inlineStr 的 `<t>` 推进、
/// `t="str"/"d"` 的 `<v>` 忽略，与 quick-xml 事件一致）；其他标签视为结构异常回退。
fn accumulate_text(content: &[u8], push_cdata: bool, out: &mut String) -> Res<()> {
    let mut pos = 0;
    while pos < content.len() {
        match memchr(b'<', &content[pos..]) {
            None => {
                let s = std::str::from_utf8(&content[pos..]).map_err(|_| FastErr)?;
                push_text_with_refs(s, out)?;
                break;
            }
            Some(i) => {
                let s = std::str::from_utf8(&content[pos..pos + i]).map_err(|_| FastErr)?;
                push_text_with_refs(s, out)?;
                let rest = &content[pos + i..];
                if rest.starts_with(b"<![CDATA[") {
                    let e = memmem::find(rest, b"]]>").ok_or(FastErr)?;
                    if push_cdata {
                        let s = std::str::from_utf8(&rest[9..e]).map_err(|_| FastErr)?;
                        push_normalized(s, out);
                    }
                    pos += i + e + 12;
                } else if rest.starts_with(b"<!--") {
                    let e = memmem::find(rest, b"-->").ok_or(FastErr)?;
                    pos += i + e + 7;
                } else {
                    return Err(FastErr);
                }
            }
        }
    }
    Ok(())
}

// ---------------------------------------------------------------- inlineStr

/// `<is>` 内容 → 字符串，语义对齐 `read_string_with_bufs`：
/// `<t>` 各段 trim（`xml:space="preserve"` 除外）+ `_x000D_` 解码后拼接；
/// 富文本 `<r>` 拼接、拼音 `<rPh>` 跳过；空 `<is>` 是合法空字符串。
fn parse_inline_str(content: &[u8]) -> Res<String> {
    let mut rich: Option<String> = None;
    let mut pos = 0;
    while let Some(i) = memchr(b'<', &content[pos..]) {
        let rest = &content[pos + i..];
        if rest.starts_with(b"<!--") {
            let e = memmem::find(rest, b"-->").ok_or(FastErr)?;
            pos += i + e + 7;
            continue;
        }
        if rest.starts_with(b"</") {
            let (name, used) = parse_end_tag(rest)?;
            if name == b"is" {
                return Ok(rich.unwrap_or_default());
            }
            // legacy 对 is 内不认识的结束标签一律忽略
            pos += i + used;
            continue;
        }
        let (tag, used) = parse_start_tag(rest)?;
        pos += i + used;
        match tag.name {
            b"r" => {
                // legacy：只有第一个 <r> 触发 rich 初始化
                if rich.is_none() {
                    rich = Some(String::new());
                }
            }
            b"rPh" => {
                if !tag.self_closing {
                    let e = memmem::find(&content[pos..], b"</rPh>").ok_or(FastErr)?;
                    pos += e + 6;
                }
            }
            b"t" => {
                let preserve = attr(tag.attrs, b"xml:space") == Some(&b"preserve"[..]);
                let mut buf = String::new();
                if !tag.self_closing {
                    let e = memmem::find(&content[pos..], b"</t>").ok_or(FastErr)?;
                    let inner = &content[pos..pos + e];
                    pos += e + 4;
                    accumulate_text(inner, true, &mut buf)?;
                }
                let value = if preserve {
                    unescape_xml(&buf).into_owned()
                } else {
                    // 只 trim ASCII 空白（NBSP 等保留）
                    unescape_xml(buf.trim_matches([' ', '\t', '\r', '\n'])).into_owned()
                };
                match &mut rich {
                    Some(s) => s.push_str(&value),
                    None => rich = Some(value),
                }
            }
            // legacy 对 is 内不认识的开始标签一律忽略（只跳过标签本身）
            _ => {}
        }
    }
    // 内容由调用者截到 </is> 为止，扫完即结束；无 <t> 的空 <is> 是合法空字符串
    Ok(rich.unwrap_or_default())
}

// ---------------------------------------------------------------- 单元格值

/// `<c>` 内容 → 值，语义对齐 `next_cell` + `read_value`：
/// 多个 `<v>`/`<is>` 时最后一个生效；`t="inlineStr"` 的冗余 `<v>` 是 Empty；
/// 纯类型（n/s/b/e/无）的 `<v>` 只取第一个 Text 事件——内容以 `&` 开头即 Empty。
fn process_cell_content(
    content: &[u8],
    t_attr: Option<&[u8]>,
    s_attr: Option<&[u8]>,
    strings: &[String],
    formats: &[CellFormat],
    is_1904: bool,
) -> Res<Data> {
    let cell_format = match s_attr {
        Some(s) => {
            let id = atoi_simd::parse::<usize, true, false>(s).unwrap_or(0);
            formats.get(id)
        }
        None => Some(&CellFormat::Other),
    };
    let mut value = Data::Empty;
    let mut pos = 0;
    while let Some(i) = memchr(b'<', &content[pos..]) {
        let rest = &content[pos + i..];
        if rest.starts_with(b"<!--") {
            let e = memmem::find(rest, b"-->").ok_or(FastErr)?;
            pos += i + e + 7;
            continue;
        }
        if rest.starts_with(b"</") {
            let (_, used) = parse_end_tag(rest)?;
            pos += i + used;
            continue;
        }
        let (tag, used) = parse_start_tag(rest)?;
        pos += i + used;
        match tag.name {
            b"f" => {
                if !tag.self_closing {
                    let e = memmem::find(&content[pos..], b"</f>").ok_or(FastErr)?;
                    pos += e + 4;
                }
            }
            b"v" => {
                let v = if tag.self_closing {
                    &[][..]
                } else {
                    let e = memmem::find(&content[pos..], b"</v>").ok_or(FastErr)?;
                    let inner = &content[pos..pos + e];
                    pos += e + 4;
                    inner
                };
                value = read_v(v, t_attr, cell_format, strings, is_1904)?;
            }
            b"is" => {
                let s = if tag.self_closing {
                    String::new()
                } else {
                    let e = memmem::find(&content[pos..], b"</is>").ok_or(FastErr)?;
                    let inner = &content[pos..pos + e];
                    pos += e + 5;
                    parse_inline_str(inner)?
                };
                value = Data::String(s);
            }
            // legacy：cell 内其他开始标签报 UnexpectedNode；回退让它报
            _ => return Err(FastErr),
        }
    }
    Ok(value)
}

/// `<v>` 的文本 → 值，语义对齐 `read_v`。
fn read_v(
    v: &[u8],
    t_attr: Option<&[u8]>,
    cell_format: Option<&CellFormat>,
    strings: &[String],
    is_1904: bool,
) -> Res<Data> {
    match t_attr {
        Some(b"s") => {
            if v.is_empty() {
                return Ok(Data::Empty);
            }
            let idx = atoi_simd::parse::<usize, true, false>(v).unwrap_or(0);
            strings
                .get(idx)
                .map(|s| Data::String(s.clone()))
                .ok_or(FastErr)
        }
        Some(b"b") => Ok(Data::Bool(v != b"0")),
        Some(b"d") => {
            let mut s = String::new();
            accumulate_text(v, false, &mut s)?;
            Ok(Data::DateTimeIso(s))
        }
        Some(b"e") => {
            let s = std::str::from_utf8(v).map_err(|_| FastErr)?;
            s.parse::<CellErrorType>()
                .map(Data::Error)
                .map_err(|_| FastErr)
        }
        Some(b"str") => {
            let mut s = String::new();
            accumulate_text(v, false, &mut s)?;
            Ok(Data::String(s))
        }
        Some(b"inlineStr") | Some(b"is") => Ok(Data::Empty),
        Some(b"n") | None => {
            if v.is_empty() {
                return Ok(Data::Empty);
            }
            // legacy 只取第一个 Text 事件：内容以 & 开头 → GeneralRef 事件 → Empty
            if v.first() == Some(&b'&') {
                return Ok(Data::Empty);
            }
            // 内容含 '<'（注释等）legacy 有定义但过于罕见：回退
            if memchr(b'<', v).is_some() {
                return Err(FastErr);
            }
            // 第一个实体之前的部分才是 legacy 读到的 Text
            let v = match memchr(b'&', v) {
                Some(i) => &v[..i],
                None => v,
            };
            let s = std::str::from_utf8(v).map_err(|_| FastErr)?;
            match fast_float2::parse::<f64, _>(s.as_bytes()) {
                Ok(n) => Ok(format_excel_f64(n, cell_format, is_1904)),
                Err(_) => match t_attr {
                    None => Ok(Data::String(s.to_string())),
                    Some(_) => Err(FastErr),
                },
            }
        }
        Some(_) => Err(FastErr), // legacy 报 CellTAttribute；回退让它报
    }
}

// ---------------------------------------------------------------- 块解析

/// 一块里解析出的行：行号属性（None = 无 r 属性）和单元格。
struct BlockOut {
    rows: Vec<Option<u32>>,
    /// (行引用, 列, 值)；Fixed 是 r 属性里的绝对行号，Local 是行下标（合并时解析）
    cells: Vec<(CellRow, u32, Data)>,
    /// 见到 `</sheetData>`，后面不用再解析
    done: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CellRow {
    Local(u32),
    Fixed(u32),
}

/// 严格解析一块（由完整 `<row>` 组成，可带前后杂项）。任何不认识的结构都报错回退。
/// 行号/列号的推进语义对齐 legacy 的 `next_cell`：行有 r 属性用之，行结束 +1；
/// 单元格有 r 属性用它的 (行, 列)，没有用顺序推进的列号。
fn parse_block(
    block: &[u8],
    strings: &[String],
    formats: &[CellFormat],
    is_1904: bool,
) -> Res<BlockOut> {
    let mut out = BlockOut {
        rows: Vec::new(),
        cells: Vec::new(),
        done: false,
    };
    let mut pos = 0;
    let mut in_row = false;
    let mut row_idx = 0usize;
    let mut cur_col: u32 = 0;

    while let Some(i) = memchr(b'<', &block[pos..]) {
        let rest = &block[pos + i..];
        // 杂项标记：注释 / 处理指令 / CDATA（legacy 对应事件一律忽略）
        if rest.starts_with(b"<!--") {
            let e = memmem::find(rest, b"-->").ok_or(FastErr)?;
            pos += i + e + 7;
            continue;
        }
        if rest.starts_with(b"<?") {
            let e = memmem::find(rest, b"?>").ok_or(FastErr)?;
            pos += i + e + 6;
            continue;
        }
        if rest.starts_with(b"<![CDATA[") {
            let e = memmem::find(rest, b"]]>").ok_or(FastErr)?;
            pos += i + e + 12;
            continue;
        }
        if rest.starts_with(b"<!") {
            return Err(FastErr); // DOCTYPE 等
        }
        if rest.starts_with(b"</") {
            let (name, used) = parse_end_tag(rest)?;
            if name == b"sheetData" {
                out.done = true;
                return Ok(out);
            }
            if name == b"row" {
                if !in_row {
                    return Err(FastErr);
                }
                in_row = false;
                cur_col = 0;
                pos += i + used;
                continue;
            }
            if name.contains(&b':') {
                return Err(FastErr); // 带前缀，legacy 的 quick-xml 能处理
            }
            pos += i + used;
            continue;
        }
        let (tag, used) = parse_start_tag(rest)?;
        pos += i + used;
        if tag.attrs.contains(&b'\'') {
            return Err(FastErr); // 单引号属性的解析没实现，回退
        }
        match tag.name {
            b"row" => {
                if in_row {
                    return Err(FastErr); // 上一个行没结束；legacy 对此有怪异语义，回退
                }
                let r = match attr(tag.attrs, b"r") {
                    Some(raw) => {
                        let s = std::str::from_utf8(raw).map_err(|_| FastErr)?;
                        Some(super::get_row(s).map_err(|_| FastErr)?)
                    }
                    None => None,
                };
                out.rows.push(r);
                row_idx = out.rows.len() - 1;
                in_row = !tag.self_closing;
            }
            b"c" => {
                if out.rows.is_empty() {
                    return Err(FastErr); // 行外单元格；legacy 对此有怪异语义，回退
                }
                let (cell_row, cell_col) = match attr(tag.attrs, b"r") {
                    Some(raw) => {
                        let s = std::str::from_utf8(raw).map_err(|_| FastErr)?;
                        let (r, c) = super::get_row_column(s).map_err(|_| FastErr)?;
                        (CellRow::Fixed(r), c)
                    }
                    None => (CellRow::Local(row_idx as u32), cur_col),
                };
                let t_attr = attr(tag.attrs, b"t");
                let s_attr = attr(tag.attrs, b"s");
                let content = if tag.self_closing {
                    &[][..]
                } else {
                    let e = memmem::find(&block[pos..], b"</c>").ok_or(FastErr)?;
                    let inner = &block[pos..pos + e];
                    pos += e + 4;
                    inner
                };
                let val = process_cell_content(content, t_attr, s_attr, strings, formats, is_1904)?;
                if !matches!(val, Data::Empty) {
                    out.cells.push((cell_row, cell_col, val));
                }
                cur_col = cell_col.saturating_add(1);
            }
            b"sheetData" => {
                if tag.self_closing {
                    out.done = true;
                    return Ok(out);
                }
            }
            // 其他已知/未知标签：只跳过标签本身，内容交给后续循环按同样的规则
            // 处理（与 legacy 忽略事件等价）；带前缀的标签回退
            _ => {
                if tag.name.contains(&b':') {
                    return Err(FastErr);
                }
            }
        }
    }
    Ok(out)
}

// ---------------------------------------------------------------- 驱动

/// 尝试用快速路径解析整个工作表；返回 None 表示应回退到 quick-xml 路径。
pub(crate) fn try_worksheet_range<RS: Read + Seek>(
    zip: &mut ZipArchive<RS>,
    cache: &HashMap<String, String>,
    path: &str,
    strings: &[String],
    formats: &[CellFormat],
    is_1904: bool,
    header_row: HeaderRow,
) -> Option<Range<Data>> {
    let entry = zip.by_name(cached_zip_path(cache, path)).ok()?;
    let mut f = BufReader::new(entry);

    // 头部探测：64 KB 内必须见到 <sheetData（带前缀的 <x:sheetData 匹配不上）
    let mut head = vec![0u8; 64 * 1024];
    let mut n = 0;
    while n < head.len() {
        match f.read(&mut head[n..]) {
            Ok(0) => break,
            Ok(k) => n += k,
            Err(_) => return None,
        }
    }
    head.truncate(n);
    memmem::find(&head, b"<sheetData")?;

    let batch = rayon::current_num_threads().clamp(2, 8);
    let (tx, rx) = mpsc::sync_channel::<Vec<u8>>(batch);
    let mut cells: Vec<Cell<Data>> = Vec::new();
    let mut failed = false;
    let mut done = false;

    // 解压句柄只在解压线程里使用；scope 期间主线程阻塞在 recv / par_iter 上，
    // 不会触碰 zip，因此独占借用不会别名
    struct AssertSend<R>(R);
    unsafe impl<R> Send for AssertSend<R> {}
    impl<R: Read> Read for AssertSend<R> {
        fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
            self.0.read(buf)
        }
    }

    let f = AssertSend(f);
    std::thread::scope(|scope| {
        scope.spawn(move || {
            let mut blocks = BlockReader::new(f, head);
            while let Ok(Some(b)) = blocks.next_block(BLOCK_SIZE) {
                if tx.send(b).is_err() {
                    break;
                }
            }
        });

        // 行号语义对齐 legacy：行有 r 属性用之；没有就上一行 + 1
        let mut next_row: u64 = 0;
        let mut row_map: Vec<u32> = Vec::new();
        'outer: while let Ok(first) = rx.recv() {
            let mut raw = vec![first];
            while raw.len() < batch {
                match rx.try_recv() {
                    Ok(b) => raw.push(b),
                    Err(_) => break,
                }
            }
            let outs: Vec<Res<BlockOut>> = raw
                .par_iter()
                .map(|b| parse_block(b, strings, formats, is_1904))
                .collect();
            for block_out in outs {
                let Ok(block_out) = block_out else {
                    failed = true;
                    break 'outer;
                };
                row_map.clear();
                row_map.reserve(block_out.rows.len());
                for r in &block_out.rows {
                    let actual = match r {
                        Some(r) => *r as u64,
                        None => next_row,
                    };
                    if actual > u32::MAX as u64 {
                        failed = true;
                        break 'outer;
                    }
                    row_map.push(actual as u32);
                    next_row = actual + 1;
                }
                for (row_ref, col, val) in block_out.cells {
                    let row = match row_ref {
                        CellRow::Fixed(r) => r,
                        CellRow::Local(idx) => match row_map.get(idx as usize) {
                            Some(row) => *row,
                            None => {
                                failed = true;
                                break 'outer;
                            }
                        },
                    };
                    cells.push(Cell::new((row, col), val));
                }
                if block_out.done {
                    done = true;
                    break 'outer;
                }
            }
        }
    });

    if failed || !done {
        return None;
    }
    Some(finish(cells, header_row))
}

/// header_row 过滤 + `Range::from_sparse`，语义与 `worksheet_range_ref` 一致。
fn finish(mut cells: Vec<Cell<Data>>, header_row: HeaderRow) -> Range<Data> {
    match header_row {
        HeaderRow::FirstNonEmptyRow => {}
        HeaderRow::Row(idx) => {
            cells.retain(|c| c.pos.0 >= idx);
            if let Some(first) = cells.first() {
                if first.pos.0 != idx {
                    let col = first.pos.1;
                    cells.insert(0, Cell::new((idx, col), Data::Empty));
                }
            }
        }
    }
    Range::from_sparse(cells)
}

#[cfg(test)]
mod tests {
    use super::super::Xlsx;
    use super::*;
    use crate::{Reader, ReaderRef};
    use std::fs::File;

    /// 对 tests/ 下所有 xlsx/xlsm fixture：fast path 能解析的，输出必须和
    /// legacy 路径（worksheet_range_ref）完全一致；处理不了的必须回退（None）。
    #[test]
    fn fast_path_matches_legacy_on_fixtures() {
        let pattern = format!(
            "{}/tests/*.xls*",
            env!("CARGO_MANIFEST_DIR").replace('\\', "/")
        );
        let mut checked = 0;
        let mut fast_used = 0;
        for entry in glob::glob(&pattern).unwrap() {
            let path = entry.unwrap();
            let file = match File::open(&path) {
                Ok(f) => f,
                Err(_) => continue,
            };
            let mut wb: Xlsx<_> = match Xlsx::new(std::io::BufReader::new(file)) {
                Ok(wb) => wb,
                Err(_) => continue, // 非 xlsx 结构的 fixture（如 xls）跳过
            };
            for (name, sheet_path) in wb.sheets.clone() {
                let legacy = {
                    let rge = wb.worksheet_range_ref(&name).unwrap();
                    Range {
                        start: rge.start,
                        end: rge.end,
                        inner: rge.inner.into_iter().map(|v| v.into()).collect(),
                    }
                };
                let fast = try_worksheet_range(
                    &mut wb.zip,
                    &wb.zip_path_cache,
                    &sheet_path,
                    &wb.strings,
                    &wb.formats,
                    wb.is_1904,
                    HeaderRow::FirstNonEmptyRow,
                );
                if let Some(fast) = fast {
                    assert_eq!(
                        fast, legacy,
                        "fast path mismatch in {path:?} sheet {name:?}"
                    );
                    fast_used += 1;
                } // None：回退路径，legacy 已验证
                checked += 1;
            }
        }
        assert!(checked > 50, "expected many fixtures, got {checked}");
        assert!(
            fast_used > 20,
            "fast path should cover most fixtures, got {fast_used}/{checked}"
        );
    }

    /// 用 HeaderRow::Row 再对一遍（过滤 + 锚点单元格语义）。
    #[test]
    fn fast_path_matches_legacy_with_header_row() {
        let pattern = format!(
            "{}/tests/*.xls*",
            env!("CARGO_MANIFEST_DIR").replace('\\', "/")
        );
        for entry in glob::glob(&pattern).unwrap() {
            let path = entry.unwrap();
            let file = match File::open(&path) {
                Ok(f) => f,
                Err(_) => continue,
            };
            let mut wb: Xlsx<_> = match Xlsx::new(std::io::BufReader::new(file)) {
                Ok(wb) => wb,
                Err(_) => continue,
            };
            for (name, sheet_path) in wb.sheets.clone() {
                wb.options.header_row = HeaderRow::Row(1);
                let legacy = {
                    let rge = wb.worksheet_range_ref(&name).unwrap();
                    Range {
                        start: rge.start,
                        end: rge.end,
                        inner: rge.inner.into_iter().map(|v| v.into()).collect(),
                    }
                };
                if let Some(fast) = try_worksheet_range(
                    &mut wb.zip,
                    &wb.zip_path_cache,
                    &sheet_path,
                    &wb.strings,
                    &wb.formats,
                    wb.is_1904,
                    HeaderRow::Row(1),
                ) {
                    assert_eq!(
                        fast, legacy,
                        "header row mismatch in {path:?} sheet {name:?}"
                    );
                }
            }
        }
    }

    fn parse(cells: &[&str], strings: &[String], formats: &[CellFormat]) -> BlockOut {
        let xml = cells.join("");
        parse_block(xml.as_bytes(), strings, formats, false).unwrap()
    }

    #[test]
    fn rows_with_gaps_and_types() {
        let out = parse(
            &[
                "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"C1\"><v>2.5</v></c></row>",
                "<row r=\"4\"><c t=\"b\"><v>1</v></c></row>",
            ],
            &[],
            &[CellFormat::Other],
        );
        assert_eq!(out.rows, vec![Some(0), Some(3)]);
        assert_eq!(
            out.cells,
            vec![
                (CellRow::Fixed(0), 0, Data::Float(1.0)),
                (CellRow::Fixed(0), 2, Data::Float(2.5)),
                (CellRow::Local(1), 0, Data::Bool(true)),
            ]
        );
    }

    #[test]
    fn rows_without_r_attr() {
        let out = parse(
            &["<row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row>"],
            &[],
            &[CellFormat::Other],
        );
        assert_eq!(out.rows, vec![None, None]);
        assert_eq!(
            out.cells,
            vec![
                (CellRow::Local(0), 0, Data::Float(1.0)),
                (CellRow::Local(0), 1, Data::Float(2.0)),
                (CellRow::Local(1), 0, Data::Float(3.0)),
            ]
        );
    }

    #[test]
    fn out_of_order_cells() {
        // 列乱序：两个都保留（from_sparse 按 pos 落位，后者覆盖前者）
        let out = parse(
            &["<row r=\"1\"><c r=\"B1\"><v>2</v></c><c r=\"A1\"><v>1</v></c></row>"],
            &[],
            &[CellFormat::Other],
        );
        assert_eq!(out.cells.len(), 2);
    }

    #[test]
    fn inline_strings() {
        let out = parse(
            &[
                "<row r=\"1\"><c t=\"inlineStr\"><is><t>  spaced  </t></is></c>",
                "<c t=\"inlineStr\"><is><t xml:space=\"preserve\">  kept  </t></is></c>",
                "<c t=\"inlineStr\"><is><r><t>foo</t></r><r><t>bar</t></r><rPh><t>拼音</t></rPh></is></c></row>",
            ],
            &[],
            &[CellFormat::Other],
        );
        assert_eq!(out.cells[0].2, Data::String("spaced".into()));
        assert_eq!(out.cells[1].2, Data::String("  kept  ".into()));
        assert_eq!(out.cells[2].2, Data::String("foobar".into()));
    }

    #[test]
    fn entities_and_shared_strings() {
        let strings = vec!["a & b".to_string(), "café".to_string()];
        let out = parse(
            &[
                "<row r=\"1\"><c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c>",
                "<c t=\"inlineStr\"><is><t>&lt;x&gt;&amp;&#65;</t></is></c>",
                "<c t=\"str\"><v>a&amp;b</v></c></row>",
            ],
            &strings,
            &[CellFormat::Other],
        );
        assert_eq!(out.cells[0].2, Data::String("a & b".into()));
        assert_eq!(out.cells[1].2, Data::String("café".into()));
        assert_eq!(out.cells[2].2, Data::String("<x>&A".into()));
        assert_eq!(out.cells[3].2, Data::String("a&b".into()));
    }

    #[test]
    fn numbers_and_errors() {
        let out = parse(
            &[
                "<row r=\"1\"><c><v>42</v></c><c><v>1e10</v></c><c t=\"e\"><v>#DIV/0!</v></c>",
                "<c><v>not-a-number</v></c></row>",
            ],
            &[],
            &[CellFormat::Other],
        );
        assert_eq!(out.cells[0].2, Data::Float(42.0));
        assert_eq!(out.cells[1].2, Data::Float(1e10));
        assert_eq!(out.cells[2].2, Data::Error(CellErrorType::Div0));
        assert_eq!(out.cells[3].2, Data::String("not-a-number".into()));
    }

    #[test]
    fn style_empty_and_formula_cells() {
        // 纯样式格、空 <v>、带公式的格都是 Empty，不出现在输出里
        let out = parse(
            &[
                "<row r=\"1\"><c r=\"A1\" s=\"1\"/><c r=\"B1\"><v></v></c>",
                "<c r=\"C1\"><f>SUM(A1:A9)</f><v>7</v></c><c t=\"s\"><v/></c></row>",
            ],
            &[],
            &[CellFormat::Other],
        );
        assert_eq!(out.cells, vec![(CellRow::Fixed(0), 2, Data::Float(7.0))]);
    }

    #[test]
    fn prefixed_or_unknown_structure_falls_back() {
        let xml =
            "<worksheet xmlns:x=\"x\"><x:row r=\"1\"><x:c><x:v>1</x:v></x:c></x:row></worksheet>";
        assert!(parse_block(xml.as_bytes(), &[], &[], false).is_err());
        // DOCTYPE
        assert!(parse_block(b"<!DOCTYPE sheet><row r=\"1\"/>", &[], &[], false).is_err());
        // sheetData 没结束：parse_block 本身成功，靠 done 标志在驱动层回退
        let out = parse_block(b"<sheetData><row r=\"1\"/>", &[], &[], false).unwrap();
        assert!(!out.done);
    }

    #[test]
    fn block_reader_cuts_on_row() {
        let data = b"<sheetData><row r=\"1\"><c><v>1</v></c></row><row r=\"2\"><c><v>2</v></c></row></sheetData>";
        let mut rd = BlockReader::new(&data[..], Vec::new());
        let mut seen = Vec::new();
        while let Some(b) = rd.next_block(16).unwrap() {
            seen.push(String::from_utf8_lossy(&b).to_string());
        }
        // 每块都对齐到 <row 起点
        for b in &seen[..seen.len() - 1] {
            assert!(b.starts_with("<row"), "block cut at {b:?}");
        }
        assert_eq!(seen.concat(), String::from_utf8_lossy(data));
    }
}
