use anyhow::{Context, Result};
use std::path::Path;

#[cfg(target_os = "windows")]
pub fn convert_doc_to_docx_word_com(src: &Path, dst: &Path) -> Result<()> {
    use windows::core::{BSTR, GUID, Interface, PCWSTR};
    use windows::Win32::System::Com::{
        CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx,
        CoUninitialize, CLSIDFromProgID, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
        DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO, IDispatch, VARIANT,
    };

    let abs_src = std::fs::canonicalize(src).unwrap_or_else(|_| src.to_path_buf());
    let src_s = abs_src.to_string_lossy().to_string();
    let dst_s = dst.to_string_lossy().to_string();

    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED)
            .ok()
            .context("Word COM 초기화 실패")?;

        let result = (|| -> Result<()> {
            let progid = wide_null("Word.Application");
            let clsid: GUID = CLSIDFromProgID(PCWSTR(progid.as_ptr()))
                .context("Word.Application COM ProgID 조회 실패")?;

            let word: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)
                .context("Word.Application 실행 실패")?;

            set_property_bool(&word, "Visible", false)?;
            set_property_i4(&word, "DisplayAlerts", 0)?;

            let documents = get_property_dispatch(&word, "Documents")?;

            // Word Documents.Open(FileName, ConfirmConversions, ReadOnly)
            // IDispatch arguments are passed in reverse order.
            let mut open_args = [
                variant_bool(true),       // ReadOnly
                variant_bool(false),      // ConfirmConversions
                variant_bstr(&src_s),     // FileName
            ];

            let doc = invoke_method_dispatch(&documents, "Open", &mut open_args)
                .context("Word에서 DOC 파일 열기 실패")?;

            clear_variants(&mut open_args);

            // wdFormatXMLDocument = 16
            // Document.SaveAs2(FileName, FileFormat)
            let mut save_args = [
                variant_i4(16),
                variant_bstr(&dst_s),
            ];

            invoke_method_void(&doc, "SaveAs2", &mut save_args)
                .context("Word SaveAs2 DOCX 저장 실패")?;

            clear_variants(&mut save_args);

            let mut close_args = [variant_bool(false)];
            let _ = invoke_method_void(&doc, "Close", &mut close_args);
            clear_variants(&mut close_args);

            let mut quit_args: [VARIANT; 0] = [];
            let _ = invoke_method_void(&word, "Quit", &mut quit_args);

            Ok(())
        })();

        CoUninitialize();

        result?;
    }

    if dst.exists() {
        Ok(())
    } else {
        anyhow::bail!(
            "DOC → DOCX 변환 실패. Word COM은 실행됐지만 결과 DOCX가 생성되지 않았습니다."
        )
    }
}

#[cfg(not(target_os = "windows"))]
pub fn convert_doc_to_docx_word_com(_src: &Path, _dst: &Path) -> Result<()> {
    anyhow::bail!("DOC 변환은 Windows에서만 지원됩니다.")
}

#[cfg(target_os = "windows")]
fn wide_null(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

#[cfg(target_os = "windows")]
unsafe fn dispid(
    obj: &windows::Win32::System::Com::IDispatch,
    name: &str,
) -> Result<i32> {
    use windows::core::{GUID, PCWSTR};

    let wide = wide_null(name);
    let names = [PCWSTR(wide.as_ptr())];
    let mut id = 0i32;

    obj.GetIDsOfNames(
        &GUID::zeroed(),
        names.as_ptr(),
        1,
        0x0409,
        &mut id,
    )
    .with_context(|| format!("COM 이름 조회 실패: {name}"))?;

    Ok(id)
}

#[cfg(target_os = "windows")]
unsafe fn invoke_raw(
    obj: &windows::Win32::System::Com::IDispatch,
    name: &str,
    flags: windows::Win32::System::Com::DISPATCH_FLAGS,
    args: &mut [windows::Win32::System::Com::VARIANT],
) -> Result<windows::Win32::System::Com::VARIANT> {
    use windows::core::GUID;
    use windows::Win32::System::Com::{DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO, VARIANT};
    use windows::Win32::System::Ole::DISPID_PROPERTYPUT;

    let id = dispid(obj, name)?;
    let mut result = VARIANT::default();
    let mut excep = EXCEPINFO::default();
    let mut arg_err = 0u32;

    let mut named = [DISPID_PROPERTYPUT];
    let is_prop_put = flags == DISPATCH_PROPERTYPUT;

    let mut params = DISPPARAMS {
        rgvarg: if args.is_empty() {
            std::ptr::null_mut()
        } else {
            args.as_mut_ptr()
        },
        rgdispidNamedArgs: if is_prop_put {
            named.as_mut_ptr()
        } else {
            std::ptr::null_mut()
        },
        cArgs: args.len() as u32,
        cNamedArgs: if is_prop_put { 1 } else { 0 },
    };

    obj.Invoke(
        id,
        &GUID::zeroed(),
        0x0409,
        flags,
        &mut params,
        Some(&mut result),
        Some(&mut excep),
        Some(&mut arg_err),
    )
    .with_context(|| format!("COM 호출 실패: {name}"))?;

    Ok(result)
}

#[cfg(target_os = "windows")]
unsafe fn invoke_method_void(
    obj: &windows::Win32::System::Com::IDispatch,
    name: &str,
    args: &mut [windows::Win32::System::Com::VARIANT],
) -> Result<()> {
    use windows::Win32::System::Com::DISPATCH_METHOD;
    use windows::Win32::System::Variant::VariantClear;

    let mut result = invoke_raw(obj, name, DISPATCH_METHOD, args)?;
    let _ = VariantClear(&mut result);
    Ok(())
}

#[cfg(target_os = "windows")]
unsafe fn invoke_method_dispatch(
    obj: &windows::Win32::System::Com::IDispatch,
    name: &str,
    args: &mut [windows::Win32::System::Com::VARIANT],
) -> Result<windows::Win32::System::Com::IDispatch> {
    use windows::core::Interface;
    use windows::Win32::System::Com::{DISPATCH_METHOD, IDispatch};
    use windows::Win32::System::Variant::{VariantClear, VT_DISPATCH};

    let mut result = invoke_raw(obj, name, DISPATCH_METHOD, args)?;

    let vt = result.Anonymous.Anonymous.vt;
    if vt != VT_DISPATCH {
        let _ = VariantClear(&mut result);
        anyhow::bail!("COM 호출 결과가 IDispatch가 아닙니다: {name}");
    }

    let p = result.Anonymous.Anonymous.Anonymous.pdispVal;
    if p.is_null() {
        let _ = VariantClear(&mut result);
        anyhow::bail!("COM 호출 결과 IDispatch 포인터가 null입니다: {name}");
    }

    let dispatch = IDispatch::from_raw(p as _);
    std::mem::forget(result);
    Ok(dispatch)
}

#[cfg(target_os = "windows")]
unsafe fn get_property_dispatch(
    obj: &windows::Win32::System::Com::IDispatch,
    name: &str,
) -> Result<windows::Win32::System::Com::IDispatch> {
    use windows::core::Interface;
    use windows::Win32::System::Com::{DISPATCH_PROPERTYGET, IDispatch};
    use windows::Win32::System::Variant::{VariantClear, VT_DISPATCH};

    let mut result = invoke_raw(obj, name, DISPATCH_PROPERTYGET, &mut [])?;

    let vt = result.Anonymous.Anonymous.vt;
    if vt != VT_DISPATCH {
        let _ = VariantClear(&mut result);
        anyhow::bail!("COM property 결과가 IDispatch가 아닙니다: {name}");
    }

    let p = result.Anonymous.Anonymous.Anonymous.pdispVal;
    if p.is_null() {
        let _ = VariantClear(&mut result);
        anyhow::bail!("COM property IDispatch 포인터가 null입니다: {name}");
    }

    let dispatch = IDispatch::from_raw(p as _);
    std::mem::forget(result);
    Ok(dispatch)
}

#[cfg(target_os = "windows")]
unsafe fn set_property_bool(
    obj: &windows::Win32::System::Com::IDispatch,
    name: &str,
    value: bool,
) -> Result<()> {
    use windows::Win32::System::Com::DISPATCH_PROPERTYPUT;
    use windows::Win32::System::Variant::VariantClear;

    let mut args = [variant_bool(value)];
    let mut result = invoke_raw(obj, name, DISPATCH_PROPERTYPUT, &mut args)?;
    let _ = VariantClear(&mut result);
    clear_variants(&mut args);
    Ok(())
}

#[cfg(target_os = "windows")]
unsafe fn set_property_i4(
    obj: &windows::Win32::System::Com::IDispatch,
    name: &str,
    value: i32,
) -> Result<()> {
    use windows::Win32::System::Com::DISPATCH_PROPERTYPUT;
    use windows::Win32::System::Variant::VariantClear;

    let mut args = [variant_i4(value)];
    let mut result = invoke_raw(obj, name, DISPATCH_PROPERTYPUT, &mut args)?;
    let _ = VariantClear(&mut result);
    clear_variants(&mut args);
    Ok(())
}

#[cfg(target_os = "windows")]
fn variant_bstr(s: &str) -> windows::Win32::System::Com::VARIANT {
    use windows::core::BSTR;
    use windows::Win32::System::Com::VARIANT;
    use windows::Win32::System::Variant::VT_BSTR;

    let mut v = VARIANT::default();
    unsafe {
        v.Anonymous.Anonymous.vt = VT_BSTR;
        v.Anonymous.Anonymous.Anonymous.bstrVal =
            std::mem::ManuallyDrop::new(BSTR::from(s));
    }
    v
}

#[cfg(target_os = "windows")]
fn variant_bool(value: bool) -> windows::Win32::System::Com::VARIANT {
    use windows::Win32::System::Com::VARIANT;
    use windows::Win32::System::Variant::VT_BOOL;

    let mut v = VARIANT::default();
    unsafe {
        v.Anonymous.Anonymous.vt = VT_BOOL;
        // VARIANT_BOOL: true = -1, false = 0
        v.Anonymous.Anonymous.Anonymous.boolVal = if value { -1 } else { 0 };
    }
    v
}

#[cfg(target_os = "windows")]
fn variant_i4(value: i32) -> windows::Win32::System::Com::VARIANT {
    use windows::Win32::System::Com::VARIANT;
    use windows::Win32::System::Variant::VT_I4;

    let mut v = VARIANT::default();
    unsafe {
        v.Anonymous.Anonymous.vt = VT_I4;
        v.Anonymous.Anonymous.Anonymous.lVal = value;
    }
    v
}

#[cfg(target_os = "windows")]
unsafe fn clear_variants(args: &mut [windows::Win32::System::Com::VARIANT]) {
    use windows::Win32::System::Variant::VariantClear;

    for v in args {
        let _ = VariantClear(v);
    }
}
