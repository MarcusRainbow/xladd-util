//! Entry point code for xladd, based on the sample C++ code
//! supplied with the Microsoft Excel12 SDK

use std::ptr;
use std::mem;
use xladd::xlcall::{LPXLOPER12, XLOPER12, xlretFailed};
use winapi::um::libloaderapi::{GetModuleHandleW, GetProcAddress};
use winapi::shared::minwindef::HMODULE;
use widestring::U16CString;
use std::ffi::CStr;
use xlauto::debug_print;

const EXCEL12ENTRYPT: &[u8] = b"MdCallBack12\0";
const XLCALL32DLL: &str = "XLCall32";
const XLCALL32ENTRYPT: &[u8] = b"GetExcel12EntryPt\0";
type EXCEL12PROC = extern "stdcall" fn(
    xlfn: ::std::os::raw::c_int, 
    count: ::std::os::raw::c_int,
    rgpxloper12: *const LPXLOPER12,
    xloper12res: LPXLOPER12) -> ::std::os::raw::c_int;
type FNGETEXCEL12ENTRYPT = extern "stdcall" fn() -> usize;

static mut XLCALL_HMODULE: HMODULE = ptr::null_mut();
static mut PEXCEL12: usize = 0;

fn fetch_excel12_entry_pt() {

    unsafe {
        if XLCALL_HMODULE.is_null() {
            let wcstr = U16CString::from_str(XLCALL32DLL).unwrap();
            XLCALL_HMODULE = GetModuleHandleW(wcstr.as_ptr());
            if !XLCALL_HMODULE.is_null() {
                debug_print("xladd-util: found hmodule for xlcall");

                let cstr = CStr::from_bytes_with_nul(XLCALL32ENTRYPT).unwrap();
                let entry_pt: usize = GetProcAddress(XLCALL_HMODULE, cstr.as_ptr()) as usize;
                if entry_pt != 0 {
                    PEXCEL12 = mem::transmute::<usize, FNGETEXCEL12ENTRYPT>(entry_pt)();
                }
            }
        }

        if PEXCEL12 == 0 {
            XLCALL_HMODULE = GetModuleHandleW(ptr::null());
            if !XLCALL_HMODULE.is_null() {
                debug_print("xladd-util: found entry pt for null");
                let cstr = CStr::from_bytes_with_nul(EXCEL12ENTRYPT).unwrap();

                PEXCEL12 = GetProcAddress(XLCALL_HMODULE, cstr.as_ptr()) as usize;
            }
        }
    }
}

pub fn excel12v(xlfn: i32, oper_res: &mut XLOPER12, opers: &[LPXLOPER12]) -> i32 {
	let md_ret;

	fetch_excel12_entry_pt();

    unsafe {
        if PEXCEL12 == 0 {
            md_ret = xlretFailed as i32;
        } else {
            let p = opers.as_ptr();
            let len = opers.len();
            md_ret = mem::transmute::<usize, EXCEL12PROC>(PEXCEL12)(xlfn, len as i32, p, oper_res);
        }
    }

	md_ret
}
