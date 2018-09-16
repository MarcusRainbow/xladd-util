use xladd::variant::Variant;
use xladd::xlcall::LPXLOPER12;

#[no_mangle]
pub extern "stdcall" fn xuVersion() -> LPXLOPER12 {
    let result = Box::new(Variant::from_str("xladd-util: version 0.1.0"));
    Box::into_raw(result) as LPXLOPER12
}
