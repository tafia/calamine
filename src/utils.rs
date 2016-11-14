//! Internal module providing handy function

/// Converts a &[u8] into a &[u32]
pub fn to_u32(s: &[u8]) -> &[u32] {
    assert!(s.len() % 4 == 0);
    unsafe { ::std::slice::from_raw_parts(s as *const [u8] as *const u32, s.len() / 4) }
}
pub fn read_slice<T>(s: &[u8]) -> T {
    unsafe { ::std::ptr::read(&s[..::std::mem::size_of::<T>()] as *const [u8] as *const T) }
}

pub fn read_u32(s: &[u8]) -> u32 {
    read_slice(s)
}

pub fn read_u16(s: &[u8]) -> u16 {
    read_slice(s)
}

pub fn read_usize(s: &[u8]) -> usize {
    read_u32(s) as usize
}
