use crate::{CellErrorType, Data};
use quick_xml::events::attributes::Attribute;
use quick_xml::events::BytesStart;
use quick_xml::name::QName;
use quick_xml::Decoder;

pub type Tag = Box<[u8]>;
pub type Value = Option<Box<[u8]>>;

pub trait PivotDataUtil {
    fn parse_item(item: (Tag, Value), decoder: &Decoder) -> Result<Data, crate::errors::Error> {
        match item.0.as_ref() {
            b"m" => Ok(Data::Empty),
            b"s" => Ok(item
                .1
                .map(|val| {
                    if let Ok(val) = decoder.decode(val.as_ref()) {
                        Data::String(val.to_string())
                    } else {
                        Data::Error(CellErrorType::GettingData)
                    }
                })
                .unwrap_or(Data::Empty)),
            b"n" => Ok(item
                .1
                .map(|val| {
                    if val.contains(&b'.') {
                        Data::Float(Self::bytes_to_f64(val.as_ref(), decoder))
                    } else {
                        Data::Int(Self::bytes_to_i64(val.as_ref(), decoder))
                    }
                })
                .unwrap_or(Data::Empty)),
            b"d" => Ok(item
                .1
                .as_ref()
                .map(|val| {
                    if let Ok(val) = decoder.decode(val) {
                        Data::DateTimeIso(val.into())
                    } else {
                        Data::Error(CellErrorType::GettingData)
                    }
                })
                .unwrap_or(Data::Empty)),
            b"b" => Ok(item
                .1
                .map(|val| {
                    Data::Bool({
                        match val.as_ref() {
                            b"0" | b"false" => false,
                            b"1" | b"true" => true,
                            _ => unreachable!("boolean tags only support W3C XML Schema"),
                        }
                    })
                })
                .unwrap_or(Data::Empty)),
            b"e" => Ok(item
                .1
                .map(|_| Data::Error(CellErrorType::Ref))
                .unwrap_or(Data::Empty)),
            _ => Err(crate::errors::Error::Msg(
                "unhandled pivot cache tag for record",
            )),
        }
    }

    fn is_item(e: &BytesStart) -> bool {
        if e.local_name().as_ref().len() > 1 {
            false
        } else {
            [b"s", b"n", b"m", b"e", b"b", b"d", b"x"]
                .into_iter()
                .any(|val| val.eq(e.local_name().as_ref()))
        }
    }

    fn data(e: &BytesStart, decoder: &Decoder) -> Result<Data, crate::errors::Error> {
        Self::parse_item(Self::byte_start_to_item(e), decoder)
    }

    fn byte_start_to_item(e: &BytesStart) -> (Tag, Value) {
        (
            Box::from(e.local_name().as_ref()),
            e.attributes().find_map(|attr| match attr {
                Ok(Attribute {
                    key: QName(b"v"),
                    value: v,
                }) => Some(Box::from(v)),
                _ => None,
            }),
        )
    }

    fn bytes_to_i64(val: &[u8], decoder: &Decoder) -> i64 {
        decoder.decode(val).unwrap().parse::<i64>().unwrap()
    }
    fn bytes_to_f64(val: &[u8], decoder: &Decoder) -> f64 {
        decoder.decode(val).unwrap().parse::<f64>().unwrap()
    }
}

pub struct PivotTableRef {
    name: String,
    sheet: String,
    records: String,
    definitions: String,
    cache_number: usize,
}

impl PivotTableRef {
    pub fn name(&self) -> &str {
        self.name.as_ref()
    }
    pub fn sheet(&self) -> &str {
        self.sheet.as_ref()
    }
    pub fn records(&self) -> &str {
        self.records.as_ref()
    }
    pub fn definitions(&self) -> &str {
        self.definitions.as_ref()
    }
    pub fn cache_number(&self) -> usize {
        self.cache_number
    }
}

#[derive(Default)]
pub struct PivotTableRefBuilder {
    name: Option<String>,
    sheet: Option<String>,
    records: Option<String>,
    definitions: Option<String>,
    cache_number: Option<usize>,
}

impl PivotTableRefBuilder {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn name(mut self, s: impl Into<String>) -> Self {
        self.name = Some(s.into());
        self
    }
    pub fn sheet(mut self, s: impl Into<String>) -> Self {
        self.sheet = Some(s.into());
        self
    }
    pub fn records(mut self, s: impl Into<String>) -> Self {
        self.records = Some(s.into());
        self
    }
    pub fn definitions(mut self, s: impl Into<String>) -> Self {
        self.definitions = Some(s.into());
        self
    }

    pub fn cache_number(mut self, n: usize) -> Self {
        self.cache_number = Some(n);
        self
    }

    pub fn build(self) -> PivotTableRef {
        PivotTableRef {
            name: self.name.unwrap(),
            sheet: self.sheet.unwrap(),
            records: self.records.unwrap(),
            definitions: self.definitions.unwrap(),
            cache_number: self.cache_number.unwrap(),
        }
    }
}
