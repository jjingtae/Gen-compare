use anyhow::{Context, Result};
use std::mem::ManuallyDrop;
use std::path::Path;

use windows::core::{w, BSTR, GUID, PCWSTR};
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize,
    DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
    DISPPARAMS, EXCEPINFO, IDispatch, CLSCTX_LOCAL_SERVER,
    COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::Variant::{
    VariantClear, VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_EMPTY, VT_I4,
};

const LCID_SYSTEM_DEFAULT: u32 = 0x0800;
const WD_FORMAT_XML_DOCUMENT: i32 = 16;
const DISPID_PROPERTYPUT: i32 = -3;

pub fn convert_doc_to_docx(src: &Path, dst: &Path) -> Result<()> {
    let src = std::fs::canonicalize(src)
        .with_context(|| format!("DOC 파일 경로 확인 실패: {}", src.display()))?;

    if let Some(parent) = dst.parent() {
        std::fs::create_dir_all(parent)
            .with_context(|| format!("출력 폴더 생성 실패: {}", parent.display()))?;
    }

    let _ = std::fs::remove_file(dst);

    unsafe {
        let _com = ComApartment::init()?;
        let word = create_word_application()?;

        let result = (|| -> Result<()> {
            word.put_property("Visible", vec![VariantValue::Bool(false)])?;
            word.put_property("DisplayAlerts", vec![VariantValue::I4(0)])?;

            let documents = word.get_property("Documents", vec![])?;

            let doc = documents.invoke_method(
                "Open",
                vec![
                    VariantValue::Bstr(src.to_string_lossy().to_string()),
                    VariantValue::Bool(false),
                    VariantValue::Bool(false),
                ],
            )?;

            // SaveAs2 시도 → 실패 시 SaveAs 폴백
            let save_result = doc.invoke_method(
                "SaveAs2",
                vec![
                    VariantValue::Bstr(dst.to_string_lossy().to_string()),
                    VariantValue::I4(WD_FORMAT_XML_DOCUMENT),
                ],
            );

            match save_result {
                Ok(_) => {}
                Err(e) if e.to_string().contains("0x80020006") => {
                    doc.invoke_method(
                        "SaveAs",
                        vec![
                            VariantValue::Bstr(dst.to_string_lossy().to_string()),
                            VariantValue::I4(WD_FORMAT_XML_DOCUMENT),
                        ],
                    )?;
                }
                Err(e) => return Err(e),
            }

            let _ = doc.invoke_method("Close", vec![VariantValue::Bool(false)]);
            let _ = word.invoke_method("Quit", vec![]);

            Ok(())
        })();

        if result.is_err() {
            let _ = word.invoke_method("Quit", vec![]);
        }

        result?;
    }

    if !dst.exists() {
        anyhow::bail!(
            "DOCX 변환은 실행됐지만 결과 파일이 생성되지 않았습니다: {}",
            dst.display()
        );
    }

    Ok(())
}

struct ComApartment;

impl ComApartment {
    unsafe fn init() -> Result<Self> {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED)
            .ok()
            .context("COM 초기화 실패")?;

        Ok(Self)
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

unsafe fn create_word_application() -> Result<ComObject> {
    let clsid = CLSIDFromProgID(w!("Word.Application"))
        .context("Word.Application CLSID 조회 실패. Microsoft Word 설치 여부를 확인하세요.")?;

    let dispatch: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)
        .context("Microsoft Word COM 인스턴스 생성 실패")?;

    Ok(ComObject { dispatch })
}

#[derive(Clone)]
struct ComObject {
    dispatch: IDispatch,
}

impl ComObject {
    unsafe fn get_dispid(&self, name: &str) -> Result<i32> {
        let wide: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
        let name_pcw = PCWSTR(wide.as_ptr());
        let names = [name_pcw];

        let mut dispid = 0i32;

        self.dispatch
            .GetIDsOfNames(
                &GUID::zeroed(),
                names.as_ptr(),
                1,
                LCID_SYSTEM_DEFAULT,
                &mut dispid,
            )
            .with_context(|| format!("COM 멤버 조회 실패: {name}"))?;

        Ok(dispid)
    }

    unsafe fn get_property(&self, name: &str, args: Vec<VariantValue>) -> Result<ComObject> {
        let result = self.invoke_raw(name, DISPATCH_PROPERTYGET, args)?;

        variant_to_dispatch(result)
            .with_context(|| format!("COM 속성 결과가 IDispatch가 아닙니다: {name}"))
    }

    unsafe fn put_property(&self, name: &str, args: Vec<VariantValue>) -> Result<()> {
        let _ = self.invoke_raw(name, DISPATCH_PROPERTYPUT, args)?;
        Ok(())
    }

    // ✅ 핵심 수정: 반환값이 IDispatch가 아닐 때 self 대신 에러 반환
    unsafe fn invoke_method(&self, name: &str, args: Vec<VariantValue>) -> Result<ComObject> {
        let result = self.invoke_raw(name, DISPATCH_METHOD, args)?;

        if variant_is_empty(&result) {
            Ok(self.clone())
        } else {
            variant_to_dispatch(result)
                .ok_or_else(|| anyhow::anyhow!("COM 메서드 반환값이 IDispatch가 아닙니다: {name}"))
        }
    }

    unsafe fn invoke_raw(
        &self,
        name: &str,
        flags: DISPATCH_FLAGS,
        args: Vec<VariantValue>,
    ) -> Result<VARIANT> {
        let dispid = self.get_dispid(name)?;

        let mut variants: Vec<VARIANT> = args
            .into_iter()
            .rev()
            .map(|value| value.into_variant())
            .collect();

        let mut named_arg = DISPID_PROPERTYPUT;

        let mut disp_params = DISPPARAMS {
            rgvarg: if variants.is_empty() {
                std::ptr::null_mut()
            } else {
                variants.as_mut_ptr()
            },
            rgdispidNamedArgs: if flags == DISPATCH_PROPERTYPUT {
                &mut named_arg
            } else {
                std::ptr::null_mut()
            },
            cArgs: variants.len() as u32,
            cNamedArgs: if flags == DISPATCH_PROPERTYPUT { 1 } else { 0 },
        };

        let mut result = VARIANT::default();
        let mut excepinfo = EXCEPINFO::default();
        let mut arg_err = 0u32;

        let invoke_result = self.dispatch.Invoke(
            dispid,
            &GUID::zeroed(),
            LCID_SYSTEM_DEFAULT,
            flags,
            &mut disp_params,
            Some(&mut result),
            Some(&mut excepinfo),
            Some(&mut arg_err),
        );

        for variant in variants.iter_mut() {
            let _ = VariantClear(variant);
        }

        invoke_result.with_context(|| format!("COM 호출 실패: {name}"))?;

        Ok(result)
    }
}

enum VariantValue {
    Bstr(String),
    Bool(bool),
    I4(i32),
}

impl VariantValue {
    unsafe fn into_variant(self) -> VARIANT {
        let mut variant = VARIANT::default();

        match self {
            VariantValue::Bstr(value) => {
                (*variant.Anonymous.Anonymous).vt = VT_BSTR;
                (*variant.Anonymous.Anonymous).Anonymous.bstrVal =
                    ManuallyDrop::new(BSTR::from(value));
            }
            VariantValue::Bool(value) => {
                (*variant.Anonymous.Anonymous).vt = VT_BOOL;
                (*variant.Anonymous.Anonymous).Anonymous.boolVal =
                    VARIANT_BOOL(if value { -1 } else { 0 });
            }
            VariantValue::I4(value) => {
                (*variant.Anonymous.Anonymous).vt = VT_I4;
                (*variant.Anonymous.Anonymous).Anonymous.lVal = value;
            }
        }

        variant
    }
}

unsafe fn variant_to_dispatch(mut variant: VARIANT) -> Option<ComObject> {
    let vt = (*variant.Anonymous.Anonymous).vt;

    if vt != VT_DISPATCH {
        let _ = VariantClear(&mut variant);
        return None;
    }

    let dispatch = ManuallyDrop::take(
        &mut (*variant.Anonymous.Anonymous).Anonymous.pdispVal,
    )?;

    Some(ComObject { dispatch })
}

unsafe fn variant_is_empty(variant: &VARIANT) -> bool {
    (*variant.Anonymous.Anonymous).vt == VT_EMPTY
}