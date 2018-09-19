//! This file contains the functions exported to Excel. My recommendation is
//! that any non-trivial logic within these functions is implemented
//! elsewhere, to keep this module clean.
//! 
//! We implement xlAutoOpen here, because it needs to register our exported
//! functions. Other xlAuto methods are exported by xladd.

use xladd::variant::Variant;
use xladd::xlcall::LPXLOPER12;
use xladd::registrator::Reg;

/// Shows version string. Note that in the Excel function wizard, this shows
/// as requiring one unnamed parameter. This is a longstanding Excel bug.
#[no_mangle]
pub extern "stdcall" fn xuVersion() -> LPXLOPER12 {
    let result = Box::new(Variant::from_str("xladd-util: version 0.1.0"));
    Box::into_raw(result) as LPXLOPER12
}

#[no_mangle]
pub extern "stdcall" fn xuGlueCols(
    a0: LPXLOPER12, a1: LPXLOPER12,
    a2: LPXLOPER12, a3: LPXLOPER12) -> LPXLOPER12 {
    let result = Box::new(Variant::concat(&[Variant::from_xloper(a0), Variant::from_xloper(a1),
        Variant::from_xloper(a2), Variant::from_xloper(a3)], true));
    Box::into_raw(result) as LPXLOPER12
}

#[no_mangle]
pub extern "stdcall" fn xuGlueRows(
    a0: LPXLOPER12, a1: LPXLOPER12,
    a2: LPXLOPER12, a3: LPXLOPER12) -> LPXLOPER12 {
    let result = Box::new(Variant::concat(&[Variant::from_xloper(a0), Variant::from_xloper(a1),
        Variant::from_xloper(a2), Variant::from_xloper(a3)], false));
    Box::into_raw(result) as LPXLOPER12
}

#[no_mangle]
pub extern "stdcall" fn xuTranspose(from: LPXLOPER12) -> LPXLOPER12 {
    let result = Box::new(Variant::transpose(&Variant::from_xloper(from)));
    Box::into_raw(result) as LPXLOPER12
}

#[no_mangle]
pub extern "stdcall" fn xlAutoOpen() -> i32 {

    let r = Reg::new();
    r.add("xuVersion", "Q$", "", "xladd-util", "Displays xladd-util version number as text.", &[]);
    r.add("xuGlueCols", "QQQQQ$", "Range, Range, Range, Range", "xladd-util", "Concatenates ranges side by side.", &[]);
    r.add("xuGlueRows", "QQQQQ$", "Range, Range, Range, Range", "xladd-util", "Concatenates ranges top to bottom.", &[]);
    r.add("xuTranspose", "QQ$", "Range", "xladd-util", "Same as TRANSPOSE, but works in any situation.", &[]);

    1
}

