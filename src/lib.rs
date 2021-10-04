#![allow(clippy::forget_copy)]

mod tabs;

use std::rc::Rc;

use bindings::Windows::Win32::Foundation::{
    self, CLASS_E_CLASSNOTAVAILABLE, E_FAIL, E_INVALIDARG, HINSTANCE, HWND, RECT, S_FALSE, S_OK,
};
use bindings::Windows::Win32::Foundation::{BOOL, E_NOTIMPL};
use bindings::Windows::Win32::Storage::StructuredStorage::IStream;
use bindings::Windows::Win32::System::Com::IOleWindow;
use bindings::Windows::Win32::UI::Shell::{
    IInputObjectSite, DBIMF_NORMAL, DBIM_ACTUAL, DBIM_BKCOLOR, DBIM_MAXSIZE, DBIM_MINSIZE,
    DBIM_MODEFLAGS, DBIM_TITLE, DESKBANDINFO,
};

use bindings::*;
use windows::*;

// {9ecce421-925a-4484-b2cf-c00b182bc32a}
const EXT_TAB_GUID: Guid = Guid::from_values(
    0x9ecce421,
    0x925a,
    0x4484,
    [0xb2, 0xcf, 0xc0, 0x0b, 0x18, 0x2b, 0xc3, 0x2a],
);

#[implement(Windows::Win32::System::Com::{IPersistStream, IObjectWithSite},  Windows::Win32::UI::Shell::{IDeskBand})]
#[derive(Default)]
struct DeskBand {
    parent_window_handle: Option<HWND>,

    p_site: Option<Rc<IUnknown>>,
    p_input_object_site: Option<Rc<IInputObjectSite>>,

    window: Option<tabs::TabBar>,
}

#[allow(non_snake_case)]
impl DeskBand {
    // IPersistStream
    pub unsafe fn GetClassID(&self) -> Result<Guid> {
        Ok(EXT_TAB_GUID)
    }

    pub unsafe fn IsDirty(&self) -> Result<()> {
        Err(S_FALSE.into())
    }

    pub unsafe fn Load(&self, _pstm: &Option<IStream>) -> Result<()> {
        Ok(())
    }

    pub unsafe fn Save(&self, _pstm: &Option<IStream>, _fcleardirty: BOOL) -> Result<()> {
        Ok(())
    }

    pub unsafe fn GetSizeMax(&self) -> Result<u64> {
        Err(E_NOTIMPL.into())
    }

    // IObjectWithSite
    pub unsafe fn SetSite(&mut self, punksite: &Option<IUnknown>) -> Result<()> {
        log::info!("Set Site");
        if self.p_site.is_some() {
            self.p_site = None;
        }

        if self.p_input_object_site.is_some() {
            self.p_input_object_site = None;
        }

        if let Some(p_unk_site) = punksite {
            self.p_input_object_site = match p_unk_site.cast() {
                Ok(n) => Some(Rc::new(n)),
                Err(_) => None,
            };

            self.parent_window_handle =
                (|| -> Result<HWND> { p_unk_site.cast::<IOleWindow>()?.GetWindow() })().ok();
            if let Some(parent) = self.parent_window_handle {
                let window = tabs::TabBar::new(parent);
                window
                    .0
                    .borrow_mut()
                    .add_tab("title\0".to_owned().as_mut_str(), 0)
                    .unwrap();
                //window.add_tab("test2");
                self.window = Some(window);
            }
        }

        log::info!("Set Site Ok");
        Ok(())
    }

    pub unsafe fn GetSite(&self, iid: *const Guid, out: *mut RawPtr) -> HRESULT {
        if let Some(site) = &self.p_input_object_site {
            return IUnknown::from(Rc::as_ref(site)).query(iid, out);
        }
        E_FAIL
    }

    // IDeskBand
    pub unsafe fn GetWindow(&self) -> Result<HWND> {
        log::info!("Get window");
        match &self.window {
            Some(window) => Ok((*window.0).borrow().handle),
            None => Err(E_FAIL.into()),
        }
    }

    pub unsafe fn ContextSensitiveHelp(&self, _: BOOL) -> Result<()> {
        Err(E_NOTIMPL.into())
    }

    pub unsafe fn ShowDW(&self, _show: BOOL) -> Result<()> {
        Ok(())
    }

    pub unsafe fn CloseDW(&mut self, _reserved: u32) -> Result<()> {
        self.p_input_object_site = None;
        Ok(())
    }
    pub unsafe fn ResizeBorderDW(
        &self,
        _prc_border: *const RECT,
        _unknown_toolbar_site: &Option<IUnknown>,
        _reserved: BOOL,
    ) -> Result<()> {
        Err(E_NOTIMPL.into())
    }

    pub unsafe fn GetBandInfo(
        &self,
        _band_id: u32,
        _view_mode: u32,
        desk_band_info_ptr: *mut DESKBANDINFO,
    ) -> Result<()> {
        if desk_band_info_ptr.is_null() {
            return Err(E_INVALIDARG.into());
        }

        log::info!("get band info id:{}, view mode:{}", _band_id, _view_mode);

        let desk_band_info = desk_band_info_ptr.as_mut().ok_or(E_INVALIDARG)?;
        if desk_band_info.dwMask & DBIM_MINSIZE != 0 {
            desk_band_info.ptMinSize.x = 200;
            desk_band_info.ptMinSize.y = 25;
        }

        if desk_band_info.dwMask & DBIM_MAXSIZE != 0 {
            desk_band_info.ptMaxSize.y = -1;
        }

        if desk_band_info.dwMask & DBIM_ACTUAL != 0 {
            desk_band_info.ptIntegral.y = -1;
        }

        if desk_band_info.dwMask & DBIM_TITLE != 0 {
            desk_band_info.dwMask &= !DBIM_TITLE;
        }

        if desk_band_info.dwMask & DBIM_MODEFLAGS != 0 {
            desk_band_info.dwModeFlags = DBIMF_NORMAL;
        }

        if desk_band_info.dwMask & DBIM_BKCOLOR != 0 {
            desk_band_info.dwMask &= !DBIM_BKCOLOR;
        }

        Ok(())
    }
}

#[implement(Windows::Win32::System::Com::IClassFactory)]
struct ClassFactory {}

#[allow(non_snake_case)]
impl ClassFactory {
    pub fn CreateInstance(
        &self,
        outer: &Option<IUnknown>,
        iid: *const Guid,
        object: *mut RawPtr,
    ) -> HRESULT {
        if outer.is_some() {
            return Foundation::CLASS_E_NOAGGREGATION;
        }

        unsafe {
            log::info!(
                "ClassFactory create instance guid:{{{:x}-...}}",
                (*iid).data1
            );
            let deskband_unknown: IUnknown = DeskBand {
                ..Default::default()
            }
            .into();
            deskband_unknown.query(iid, object)
        }
    }

    pub unsafe fn LockServer(&self, _flock: BOOL) -> Result<()> {
        Ok(())
    }
}

// Dll stuff
#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllMain(
    _instance: HINSTANCE,
    dw_reason: u32,
    _lpv_reserved: RawPtr,
) -> BOOL {
    if dw_reason == 1 {
        //DLL_PROCESS_ATTACH
        fern::Dispatch::new()
            .level(log::LevelFilter::Debug)
            .format(|out, message, record| {
                out.finish(format_args!(
                    "{}[{}][{}] {}",
                    chrono::Local::now().format("[%Y-%m-%d][%H:%M:%S]"),
                    record.target(),
                    record.level(),
                    message
                ))
            })
            .chain(fern::log_file("D:/Dev/Tabs/output.log").unwrap())
            .apply()
            .unwrap();
        log::info!("Attached");
        std::panic::set_hook(Box::new(|info| log::error!("PANIC ! {:?}", info)));

        // Make this safe at some point
        unsafe {
            tabs::DLL_INSTANCE = Some(_instance);
        }
    }
    true.into()
}

/// # Safety
#[no_mangle]
#[allow(non_snake_case)]
pub unsafe extern "stdcall" fn DllGetClassObject(
    rclsid: *const Guid,
    iid: *const Guid,
    object: *mut RawPtr,
) -> HRESULT {
    if EXT_TAB_GUID == *rclsid {
        log::info!("Dll Got ClassObject");
        let unknown: IUnknown = ClassFactory {}.into();
        return unknown.query(iid, object);
    }
    CLASS_E_CLASSNOTAVAILABLE
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "stdcall" fn DllCanUnloadNow() -> HRESULT {
    S_OK
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "stdcall" fn DllRegisterServer() -> HRESULT {
    S_OK
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "stdcall" fn DllUnregisterServer() -> HRESULT {
    S_OK
}
