#[no_mangle]
pub extern "stdcall" fn FFIDivide(x: i32, y: i32) -> i32 {
    x / y
}

#[no_mangle]
pub extern "stdcall" fn FFIMultiply(x: i32, y: i32) -> i32 {
    x * y
}

#[no_mangle]
pub extern "stdcall" fn FFISubtract(x: i32, y: i32) -> i32 {
    x - y
}

#[no_mangle]
pub extern "stdcall" fn FFISum(x: i32, y: i32) -> i32 {
    x + y
}
