//! Functions that are exported from the xll and invoked by Excel

use xladd::xlcall::{LPXLOPER12, xlerrValue, xlCoerce, xlGetName, xlfRegister };
use xladd::entrypoint::{excel12};
use std::ffi::CString;
use winapi::um::debugapi::OutputDebugStringA;
use xladd::variant::Variant;

pub fn debug_print(message: &str) {
    let cstr = CString::new(message).unwrap();
    unsafe { OutputDebugStringA(cstr.as_ptr()) };
}

#[no_mangle]
pub extern "stdcall" fn xlAutoOpen() -> i32 {

    let name = excel12(xlGetName, &mut []);
    debug_print(&format!("xladd-util loaded from: {}", name));

    let mut opers = vec![
        name.clone(),                    // dll name
        Variant::from_str("xuVersion"),  // export name
        Variant::from_str("Q$"),         // return XLOPER, thread safe
        Variant::from_str("xuVersion"),  // name in Excel
        Variant::from_str(""),           // help string for args (there are none here)
        Variant::from_int(1),            // type 1 means useable anywhere (spreadsheet or macro code)
        Variant::from_str("xladd-util"), // category (can be an existing category or a new one as here)
        Variant::missing(),              // no shortcut
        Variant::missing(),              // no help URL
        Variant::from_str("xladd-util version"),    // help
        // would be followed by argument help strings if there were any
    ];

    excel12(xlfRegister, &mut opers);
    1
}

#[no_mangle]
pub extern "stdcall" fn xlAddInManagerInfo12(px_action: LPXLOPER12) -> LPXLOPER12 {

    // This code coerces the passed-in value to an integer. This is how the
    // code determines what is being requested. If it receives a 1, it returns a
    // string representing the long name. If it receives anything else, it
    // returns a #VALUE! error.

    let action = excel12(xlCoerce, &mut vec![Variant::from_xloper(px_action), Variant::from_int(0)]);
    let result = match action.as_i32() {
        Some(i) if i == 1 => {
            Box::new(Variant::from_str("xladd-util: Simple excel utilities written in Rust"))
        },
        _ => {
            Box::new(Variant::from_err(xlerrValue))
        }
    };
    Box::into_raw(result) as LPXLOPER12
}

#[no_mangle]
pub extern "stdcall" fn xlAutoFree12(px_free: LPXLOPER12) {
    // take ownership of this xloper. Then when our xloper goes
    // out of scope, its drop method will free any resources.
    unsafe { Box::from_raw(px_free) };
}
