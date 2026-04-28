//! Direct Microsoft Word COM automation for DOC -> DOCX conversion.
//! This avoids VBS/PowerShell scripts. Requires the `windows` crate with COM/OLE features.

use anyhow::{Context, Result};
use std::path::{Path, PathBuf};

#[cfg(windows)]
pub fn convert_doc_to_docx(src: &Path, dst: &Path) -> Result<()> {
    use windows::core::{BSTR, GUID, PCWSTR, PWSTR};
    use windows::Win32::System::Com::{
        CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_LOCAL_SERVER,
        COINIT_APARTMENTTHREADED, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPPARAMS,
        EXCEPINFO, IDispatch, CLSIDFromProgID,
    };
    use windows::Win32::System::Ole::{
        VariantClear, DISPID_PROPERTYPUT, DISPATCH_PROPERTYPUT, VARIANT,
        VT_BOOL, VT_BSTR, VT_DISPATCH, VT_I4,
    };

    struct ComGuard;
    impl Drop for ComGuard {
        fn drop(&mut self) { unsafe { CoUninitialize(); } }
    }

    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED)
            .context("initialize COM for Word automation")?;
        let _guard = ComGuard;

        let progid = wide_null("Word.Application");
        let mut clsid = GUID::zeroed();
        CLSIDFromProgID(PCWSTR(progid.as_ptr()), &mut clsid)
            .context("resolve Word.Application COM class")?;

        let word: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)
            .context("create Word.Application COM object")?;

        // Word.Visible = False, Word.DisplayAlerts = 0
        set_property_bool(&word, "Visible", false)?;
        set_property_i4(&word, "DisplayAlerts", 0)?;

        let documents = get_property_dispatch(&word, "Documents")?;
        let abs_src = std::fs::canonicalize(src).unwrap_or_else(|_| src.to_path_buf());
        if let Some(parent) = dst.parent() { std::fs::create_dir_all(parent)?; }

        let opened = invoke_method_dispatch(
            &documents,
            "Open",
            &mut [
                variant_bstr(&abs_src.to_string_lossy()),
                variant_bool(false), // ConfirmConversions
                variant_bool(false), // ReadOnly: false so SaveAs2 can convert reliably
            ],
        ).context("open DOC in Word")?;

        // FileFormat 16 = wdFormatXMLDocument (.docx)
        invoke_method_void(
            &opened,
            "SaveAs2",
            &mut [variant_bstr(&dst.to_string_lossy()), variant_i4(16)],
        ).context("save DOC as DOCX via Word")?;

        let _ = invoke_method_void(&opened, "Close", &mut [variant_bool(false)]);
        let _ = invoke_method_void(&word, "Quit", &mut []);
    }

    if dst.exists() { Ok(()) } else { anyhow::bail!("Word COM completed but DOCX was not created") }
}

#[cfg(not(windows))]
pub fn convert_doc_to_docx(_src: &Path, _dst: &Path) -> Result<()> {
    anyhow::bail!("DOC conversion requires Windows + Microsoft Word")
}

#[cfg(windows)]
fn wide_null(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

#[cfg(windows)]
unsafe fn dispid(obj: &windows::Win32::System::Com::IDispatch, name: &str) -> Result<i32> {
    use windows::core::{PCWSTR, PWSTR};
    let mut wide = wide_null(name);
    let mut name_ptr = PWSTR(wide.as_mut_ptr());
    let mut id = 0i32;
    obj.GetIDsOfNames(&windows::core::GUID::zeroed(), &mut name_ptr, 1, 0x0409, &mut id)
        .with_context(|| format!("resolve COM member {name}"))?;
    Ok(id)
}

#[cfg(windows)]
unsafe fn invoke_raw(
    obj: &windows::Win32::System::Com::IDispatch,
    name: &str,
    flags: u16,
    args: &mut [windows::Win32::System::Ole::VARIANT],
) -> Result<windows::Win32::System::Ole::VARIANT> {
    use windows::Win32::System::Com::{DISPPARAMS, EXCEPINFO};
    use windows::Win32::System::Ole::{DISPID_PROPERTYPUT, VARIANT};

    let id = dispid(obj, name)?;
    // COM expects arguments in reverse order.
    args.reverse();
    let mut result = VARIANT::default();
    let mut excep = EXCEPINFO::default();
    let mut arg_err = 0u32;
    let mut named = [DISPID_PROPERTYPUT];
    let mut params = DISPPARAMS {
        rgvarg: if args.is_empty() { std::ptr::null_mut() } else { args.as_mut_ptr() },
        rgdispidNamedArgs: if flags == windows::Win32::System::Ole::DISPATCH_PROPERTYPUT.0 as u16 { named.as_mut_ptr() } else { std::ptr::null_mut() },
        cArgs: args.len() as u32,
        cNamedArgs: if flags == windows::Win32::System::Ole::DISPATCH_PROPERTYPUT.0 as u16 { 1 } else { 0 },
    };
    obj.Invoke(id, &windows::core::GUID::zeroed(), 0x0409, flags, &mut params, Some(&mut result), Some(&mut excep), Some(&mut arg_err))
        .with_context(|| format!("invoke COM member {name}"))?;
    Ok(result)
}

#[cfg(windows)]
unsafe fn invoke_method_void(obj: &windows::Win32::System::Com::IDispatch, name: &str, args: &mut [windows::Win32::System::Ole::VARIANT]) -> Result<()> {
    let mut result = invoke_raw(obj, name, windows::Win32::System::Com::DISPATCH_METHOD.0 as u16, args)?;
    let _ = windows::Win32::System::Ole::VariantClear(&mut result);
    Ok(())
}

#[cfg(windows)]
unsafe fn invoke_method_dispatch(obj: &windows::Win32::System::Com::IDispatch, name: &str, args: &mut [windows::Win32::System::Ole::VARIANT]) -> Result<windows::Win32::System::Com::IDispatch> {
    let mut result = invoke_raw(obj, name, windows::Win32::System::Com::DISPATCH_METHOD.0 as u16, args)?;
    let vt = unsafe { result.Anonymous.Anonymous.vt };
    if vt.0 != windows::Win32::System::Ole::VT_DISPATCH.0 {
        let _ = windows::Win32::System::Ole::VariantClear(&mut result);
        anyhow::bail!("COM method {name} did not return IDispatch");
    }
    let p = unsafe { result.Anonymous.Anonymous.Anonymous.pdispVal };
    if p.is_null() { anyhow::bail!("COM method {name} returned null IDispatch"); }
    let dispatch = unsafe { windows::Win32::System::Com::IDispatch::from_raw(p) };
    // Do not VariantClear result after from_raw, because ownership moved into IDispatch.
    Ok(dispatch)
}

#[cfg(windows)]
unsafe fn get_property_dispatch(obj: &windows::Win32::System::Com::IDispatch, name: &str) -> Result<windows::Win32::System::Com::IDispatch> {
    let mut result = invoke_raw(obj, name, windows::Win32::System::Com::DISPATCH_PROPERTYGET.0 as u16, &mut [])?;
    let vt = unsafe { result.Anonymous.Anonymous.vt };
    if vt.0 != windows::Win32::System::Ole::VT_DISPATCH.0 {
        let _ = windows::Win32::System::Ole::VariantClear(&mut result);
        anyhow::bail!("COM property {name} did not return IDispatch");
    }
    let p = unsafe { result.Anonymous.Anonymous.Anonymous.pdispVal };
    if p.is_null() { anyhow::bail!("COM property {name} returned null IDispatch"); }
    Ok(unsafe { windows::Win32::System::Com::IDispatch::from_raw(p) })
}

#[cfg(windows)]
unsafe fn set_property_bool(obj: &windows::Win32::System::Com::IDispatch, name: &str, value: bool) -> Result<()> {
    let mut args = [variant_bool(value)];
    let mut result = invoke_raw(obj, name, windows::Win32::System::Ole::DISPATCH_PROPERTYPUT.0 as u16, &mut args)?;
    let _ = windows::Win32::System::Ole::VariantClear(&mut result);
    Ok(())
}

#[cfg(windows)]
unsafe fn set_property_i4(obj: &windows::Win32::System::Com::IDispatch, name: &str, value: i32) -> Result<()> {
    let mut args = [variant_i4(value)];
    let mut result = invoke_raw(obj, name, windows::Win32::System::Ole::DISPATCH_PROPERTYPUT.0 as u16, &mut args)?;
    let _ = windows::Win32::System::Ole::VariantClear(&mut result);
    Ok(())
}

#[cfg(windows)]
fn variant_bstr(s: &str) -> windows::Win32::System::Ole::VARIANT {
    use windows::Win32::System::Ole::{VARIANT, VT_BSTR};
    let mut v = VARIANT::default();
    unsafe {
        v.Anonymous.Anonymous.vt = VT_BSTR;
        v.Anonymous.Anonymous.Anonymous.bstrVal = std::mem::ManuallyDrop::new(windows::core::BSTR::from(s));
    }
    v
}

#[cfg(windows)]
fn variant_bool(value: bool) -> windows::Win32::System::Ole::VARIANT {
    use windows::Win32::System::Ole::{VARIANT, VT_BOOL};
    let mut v = VARIANT::default();
    unsafe {
        v.Anonymous.Anonymous.vt = VT_BOOL;
        v.Anonymous.Anonymous.Anonymous.boolVal = if value { -1i16 } else { 0i16 };
    }
    v
}

#[cfg(windows)]
fn variant_i4(value: i32) -> windows::Win32::System::Ole::VARIANT {
    use windows::Win32::System::Ole::{VARIANT, VT_I4};
    let mut v = VARIANT::default();
    unsafe {
        v.Anonymous.Anonymous.vt = VT_I4;
        v.Anonymous.Anonymous.Anonymous.lVal = value;
    }
    v
}
