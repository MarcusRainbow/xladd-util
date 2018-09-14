pub mod entrypoint;
pub mod xlauto;

extern crate xladd;
extern crate winapi;
extern crate widestring;

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
