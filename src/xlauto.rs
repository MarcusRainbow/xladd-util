//! Functions that are exported from the xll and invoked by Excel

use xladd::xlcall::xloper12__bindgen_ty_1;
use xladd::xlcall::{XLOPER12, LPXLOPER12, xltypeNil, xltypeInt, xltypeStr, xltypeErr, xltypeMissing,
    xlerrValue, xlCoerce, xlGetName, xlfRegister };
use entrypoint::excel12v;
use std::mem;
use std::ffi::CString;
use std::slice;
use winapi::um::debugapi::OutputDebugStringA;

pub fn debug_print(message: &str) {
    let cstr = CString::new(message).unwrap();
    unsafe { OutputDebugStringA(cstr.as_ptr()) };
}

pub fn xloper_from_str(text: &str) -> XLOPER12 {
    let mut wstr : Vec<u16> = text.encode_utf16().collect();
    let len = wstr.len() as u16;
    wstr.insert(0, len);
    let raw_chars: *mut u16 = wstr.as_mut_ptr();
    mem::forget(wstr);
    XLOPER12 { xltype : xltypeStr, val: xloper12__bindgen_ty_1 { str: raw_chars } }
}

#[no_mangle]
pub extern "stdcall" fn xlAutoOpen() -> i32 {

    let mut name: XLOPER12 = XLOPER12 { xltype : xltypeNil, val: xloper12__bindgen_ty_1 { w: 0 } };

    excel12v(xlGetName as i32, &mut name, &[]);

    let cstr: *const u16 = unsafe { name.val.str };
    let cstr_len = unsafe { *cstr.offset(0) as usize };
    let cstr_slice = unsafe { slice::from_raw_parts(cstr.offset(1), cstr_len) };
    let rust_name = String::from_utf16(cstr_slice).unwrap();

    debug_print(&format!("xladd-util loaded from: {}", rust_name));

    let mut opers: Vec<XLOPER12> = Vec::with_capacity(20);  // don't reallocate
    let mut popers: Vec<LPXLOPER12> = Vec::new();

    popers.push(&mut name);
    opers.push(xloper_from_str("xuVersion"));
    opers.push(xloper_from_str("Q$"));   // return XLOPER, thread safe
    opers.push(xloper_from_str("xuVersion"));
    opers.push(xloper_from_str(""));    // no args
    opers.push(XLOPER12 { xltype : xltypeInt, val: xloper12__bindgen_ty_1 { w: 1 } }); // type 1
    opers.push(xloper_from_str("Information"));    // category
    opers.push(XLOPER12 { xltype : xltypeMissing, val: xloper12__bindgen_ty_1 { w: 0 } }); // no shortcut
    opers.push(XLOPER12 { xltype : xltypeMissing, val: xloper12__bindgen_ty_1 { w: 0 } }); // no help URL
    opers.push(xloper_from_str("xladd-util version"));    // help
    // would be followed by argument help strings if there were any

    for oper in opers.iter_mut() {
        popers.push(oper);
    }

    let mut registered: XLOPER12 = XLOPER12 { xltype : xltypeNil, val: xloper12__bindgen_ty_1 { w: 0 } };
    excel12v(xlfRegister as i32, &mut registered, &popers);

    1
}

#[no_mangle]
pub extern "stdcall" fn xlAddInManagerInfo12(px_action: LPXLOPER12) -> LPXLOPER12 {
    static mut X_INFO: XLOPER12 = XLOPER12 { xltype : xltypeNil, val: xloper12__bindgen_ty_1 { w: 0 } };
    let mut x_int_action: XLOPER12 = XLOPER12 { xltype : xltypeNil, val: xloper12__bindgen_ty_1 { w: 0 } };
/*
** This code coerces the passed-in value to an integer. This is how the
** code determines what is being requested. If it receives a 1, it returns a
** string representing the long name. If it receives anything else, it
** returns a #VALUE! error.
*/
    let mut type_int = XLOPER12 { xltype: xltypeInt, val: xloper12__bindgen_ty_1 { w: xltypeInt as i32 } };
    let args = vec![px_action, &mut type_int];
    excel12v(xlCoerce as i32, &mut x_int_action, &args);
    if unsafe { x_int_action.val.w == 1 } {
        unsafe { X_INFO.xltype = xltypeStr };

        let s = "xladd-util: Simple excel utilities written in Rust";
        let mut wstr : Vec<u16> = s.encode_utf16().collect();
        let len = wstr.len() as u16;
        wstr.insert(0, len);
        unsafe { X_INFO.val.str = wstr.as_mut_ptr() };
        mem::forget(wstr);
    }
    else {
        unsafe {
            X_INFO.xltype = xltypeErr;
            X_INFO.val.err = xlerrValue as i32;
        }
    }
// Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
// for UDFs declared as thread safe. Use alternate memory allocation mechanisms.
    unsafe { &mut X_INFO }
}

#[no_mangle]
pub extern "stdcall" fn xlAutoFree12(_px_free: LPXLOPER12) {
    // do nothing for now -- this will leak
}

#[no_mangle]
pub extern "stdcall" fn xuVersion() -> LPXLOPER12 {
    static mut VERSION: XLOPER12 = XLOPER12 { xltype : xltypeNil, val: xloper12__bindgen_ty_1 { w: 0 } };

    unsafe { VERSION.xltype = xltypeStr };

    let s = "xladd-utils version 0.1.0";
    let mut wstr : Vec<u16> = s.encode_utf16().collect();
    let len = wstr.len() as u16;
    wstr.insert(0, len);
    unsafe { VERSION.val.str = wstr.as_mut_ptr() };
    mem::forget(wstr);

// Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
// for UDFs declared as thread safe. Use alternate memory allocation mechanisms.
    unsafe { &mut VERSION }
}
