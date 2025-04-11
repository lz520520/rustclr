use super::WinStr;
use crate::error::ClrError;
use crate::Result;
use std::{
    ffi::c_void, 
    ptr::{copy_nonoverlapping, null_mut}
};
use windows_sys::Win32::{
    Foundation::{
        SysFreeString, VARIANT_FALSE, 
        VARIANT_TRUE
    }, 
    System::{
        Com::{SAFEARRAY, SAFEARRAYBOUND}, 
        Ole::{
            SafeArrayAccessData, SafeArrayCreate, 
            SafeArrayCreateVector, SafeArrayPutElement, 
            SafeArrayUnaccessData
        }, 
        Variant::{
            VARIANT, VT_ARRAY, VT_BSTR, VT_BOOL, 
            VT_I4, VT_UI1, VT_VARIANT,
        } 
    }
};


/// Trait to convert various Rust types to Windows COM-compatible `VARIANT` types.
/// 
/// This trait is implemented for common Rust types like `String`, `&str`, `bool`, and `i32`.
pub trait Variant {
    /// Converts the Rust type to a `VARIANT`.
    ///
    /// # Returns
    ///
    /// * The corresponding `VARIANT` structure for the implementing type.
    fn to_variant(&self) -> VARIANT;
    
    /// Returns the `u16` representing the VARIANT type.
    ///
    /// # Returns
    ///
    /// * The type ID for the VARIANT.
    fn var_type() -> u16;
}

impl Variant for String {
    /// Converts a `String` to a BSTR-based `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let bstr = self.to_bstr();
        let mut variant = unsafe { std::mem::zeroed::<VARIANT>() }; 
        variant.Anonymous.Anonymous.vt = Self::var_type();
        variant.Anonymous.Anonymous.Anonymous.bstrVal = bstr;

        variant
    }

    /// Returns the VARIANT type ID for BSTRs.
    fn var_type() -> u16 {
        VT_BSTR
    }
}

impl Variant for &str {
    /// Converts a `&str` to a BSTR-based `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let bstr = self.to_bstr();
        let mut variant = unsafe { std::mem::zeroed::<VARIANT>() }; 
        variant.Anonymous.Anonymous.vt = Self::var_type();
        variant.Anonymous.Anonymous.Anonymous.bstrVal = bstr;

        variant
    }

    /// Returns the VARIANT type ID for BSTRs.
    fn var_type() -> u16 {
        VT_BSTR
    }
}

impl Variant for bool {
    /// Converts a `bool` to a boolean `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let mut variant = unsafe { std::mem::zeroed::<VARIANT>() };
        variant.Anonymous.Anonymous.vt = Self::var_type();
        variant.Anonymous.Anonymous.Anonymous.boolVal = if *self {
            VARIANT_TRUE
        } else {
            VARIANT_FALSE
        };

        variant
    }

    /// Returns the VARIANT type ID for booleans.
    fn var_type() -> u16 {
        VT_BOOL
    }
}

impl Variant for i32 {
    /// Converts an `i32` to an integer `VARIANT`.
    fn to_variant(&self) -> VARIANT {
        let mut variant = unsafe { std::mem::zeroed::<VARIANT>() };
        variant.Anonymous.Anonymous.vt = Self::var_type();
        variant.Anonymous.Anonymous.Anonymous.lVal = *self;

        variant
    }

    /// Returns the VARIANT type ID for integers.
    fn var_type() -> u16 {
        VT_I4
    }
}

/// Creates a `SAFEARRAY` from a vector of elements implementing the `Variant` trait.
///
/// # Arguments
///
/// * `args` - A vector of elements implementing the `Variant` trait.
///
/// # Returns
///
/// * `Ok(*mut SAFEARRAY)` - The created `SAFEARRAY`.
/// * `Err(ClrError)` - If the creation or element insertion into the `SAFEARRAY` fails.
pub fn create_safe_array_args<T: Variant>(args: Vec<T>) -> Result<*mut SAFEARRAY> {
    unsafe {
        let vartype = T::var_type();
        let psa = SafeArrayCreateVector(vartype, 0, args.len() as u32);
        if psa.is_null() {
            return Err(ClrError::NullPointerError("SafeArrayCreateVector"));
        }
        
        for (i, arg) in args.iter().enumerate() {
            let variant = arg.to_variant();
            let index = i as i32;
            let value_ptr = match vartype {
                VT_BOOL => &variant.Anonymous.Anonymous.Anonymous.boolVal as *const _ as *const c_void,
                VT_I4 => &variant.Anonymous.Anonymous.Anonymous.lVal as *const _ as *const c_void,
                VT_BSTR => variant.Anonymous.Anonymous.Anonymous.bstrVal as *const c_void,
                _ => return Err(ClrError::VariantUnsupported),
            };

            let hr = SafeArrayPutElement(psa, &index, value_ptr);
            if hr != 0 {
                return Err(ClrError::ApiError("SafeArrayPutElement", hr));
            }

            if vartype == VT_BSTR {
                SysFreeString(variant.Anonymous.Anonymous.Anonymous.bstrVal);
            }
        }
        
        let args = SafeArrayCreateVector(VT_VARIANT, 0, 1);
        let mut var_array = std::mem::zeroed::<VARIANT>(); 
        var_array.Anonymous.Anonymous.vt = VT_ARRAY | vartype;
        var_array.Anonymous.Anonymous.Anonymous.parray = psa;

        let index = 0;
        let hr = SafeArrayPutElement(
            args, 
            &index, 
            &mut var_array as *const VARIANT as *const c_void
        );
        if hr != 0 {
            return Err(ClrError::ApiError("SafeArrayPutElement (2)", hr));
        }
    
        Ok(args)
    }
}

/// Creates a `SAFEARRAY` from a vector of `VARIANT` elements.
///
/// # Arguments
///
/// * `args` - A vector of `VARIANT` elements.
///
/// # Returns
///
/// * `Ok(*mut SAFEARRAY)` - The created `SAFEARRAY`.
/// * `Err(ClrError)` - If the creation or element insertion into the `SAFEARRAY` fails.
pub fn create_safe_args(args: Vec<VARIANT>) -> Result<*mut SAFEARRAY> {
    unsafe {       
        let arg = SafeArrayCreateVector(VT_VARIANT, 0, args.len() as u32);
        for (i, var) in args.iter().enumerate() {
            let index = i as i32;
            let mut variant = *var;
            let hr = SafeArrayPutElement(
                arg, 
                &index, 
                &mut variant as *const VARIANT as *const c_void
            );
            if hr != 0 {
                return Err(ClrError::ApiError("SafeArrayPutElement", hr));
            }
        }

        Ok(arg)
    }
}

/// Creates a `SAFEARRAY` from a byte buffer for loading assemblies.
///
/// # Arguments
///
/// * `data` - A byte slice representing the data.
///
/// # Returns
///
/// * `Ok(*mut SAFEARRAY)` - The created `SAFEARRAY`.
/// * `Err(ClrError)` - If the creation or data copying into the `SAFEARRAY` fails.
pub fn create_safe_array_buffer(data: &[u8]) -> Result<*mut SAFEARRAY> {
    let len: u32 = data.len() as u32;
    let bounds = SAFEARRAYBOUND {
        cElements: data.len() as _,
        lLbound: 0,
    };

    unsafe {
        let sa = SafeArrayCreate(VT_UI1, 1, &bounds);
        if sa.is_null() {
            return Err(ClrError::NullPointerError("SafeArrayCreate"));
        }
    
        let mut p_data = null_mut();
        let mut hr = SafeArrayAccessData(sa, &mut p_data);
        if hr != 0 {
            return Err(ClrError::ApiError("SafeArrayAccessData", hr));
        }
    
        copy_nonoverlapping(data.as_ptr(), p_data as *mut u8, len as usize);
        hr = SafeArrayUnaccessData(sa);
        if hr != 0 {
            return Err(ClrError::ApiError("SafeArrayUnaccessData", hr));
        }
    
        Ok(sa)
    }
}
