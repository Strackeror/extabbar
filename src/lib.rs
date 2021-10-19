#![allow(clippy::forget_copy)]

mod detour;
mod tabs;

use std::ffi::c_void;
use std::rc::{Rc, Weak};

use bindings::Windows::Win32::System::LibraryLoader::DisableThreadLibraryCalls;
use bindings::Windows::Win32::System::SystemServices::IServiceProvider;
use bindings::Windows::Win32::System::WindowsProgramming::{
    DWebBrowserEvents2, IWebBrowser2, IWebBrowserApp,
};
use bindings::Windows::Win32::UI::WindowsAndMessaging::{
    DestroyWindow, ShowWindow, SW_HIDE, SW_SHOW,
};
use bindings::*;
use windows::*;

use bindings::Windows::Win32::Foundation::*;
use bindings::Windows::Win32::System::Com::{IConnectionPointContainer, IOleWindow};
use bindings::Windows::Win32::System::OleAutomation::*;
use bindings::Windows::Win32::UI::Shell::*;

use crate::tabs::tab_bar::DLL_INSTANCE;

// {9ecce421-925a-4484-b2cf-c00b182bc32a}
const EXT_TAB_GUID: Guid = Guid::from_values(
    0x9ecce421,
    0x925a,
    0x4484,
    [0xb2, 0xcf, 0xc0, 0x0b, 0x18, 0x2b, 0xc3, 0x2a],
);

static mut DLL_LOCK: i32 = 0;

pub const BROWSE_OBJECT_MESSAGE: &str = "extabbar_BrowseObject";
pub const SHOW_WINDOW_MESSAGE: &str = "extabbar_ShowWindow";

#[implement(Windows::Win32::System::WindowsProgramming::DWebBrowserEvents2)]
#[derive(Clone)]
struct BrowserEventHandler {
    tab_bar: Weak<tabs::tab_bar::TabBar>,
    browser: IShellBrowser,
}

pub fn get_current_folder_pidl(browser: &IShellBrowser) -> Result<*mut ITEMIDLIST> {
    unsafe {
        let folder_view: IFolderView = browser.QueryActiveShellView()?.cast()?;
        let folder = folder_view.GetFolder::<IPersistFolder2>()?;
        let folder_pidl = folder.GetCurFolder();
        if folder_pidl.is_err() {
            log::error!("Could not get pidl for current path");
        }
        folder_pidl
    }
}

fn query_service_provider<T>(service_provider: &IServiceProvider) -> Result<T>
where
    T: Interface,
    T: Abi<Abi = *mut std::ffi::c_void>,
{
    let guid = T::IID;
    let mut ptr: *mut c_void = std::ptr::null_mut();
    unsafe {
        service_provider.QueryService(
            std::ptr::addr_of!(guid),
            std::ptr::addr_of!(guid),
            std::ptr::addr_of_mut!(ptr),
        )?;
        T::from_abi(ptr)
    }
}

#[allow(non_snake_case, unused_variables)]
impl BrowserEventHandler {
    pub unsafe fn GetTypeInfoCount(&self) -> Result<u32> {
        Ok(0)
    }

    pub unsafe fn GetTypeInfo(&self, itinfo: u32, lcid: u32) -> Result<ITypeInfo> {
        Err(E_NOTIMPL.into())
    }

    pub unsafe fn GetIDsOfNames(
        &self,
        riid: *const Guid,
        rgsznames: *const PWSTR,
        cnames: u32,
        lcid: u32,
        rgdispid: *mut i32,
    ) -> Result<()> {
        log::info!("rgsznames:{:?} lcid:{:?}", *rgsznames, lcid);
        Err(DISP_E_UNKNOWNNAME.into())
    }

    fn NavigateComplete(&self, params: &[VARIANT]) -> Result<VARIANT> {
        let path = get_current_folder_pidl(&self.browser);
        self.tab_bar.upgrade().unwrap().navigated(path.ok())?;

        Ok(Default::default())
    }

    unsafe fn NewWindow(&mut self, params: &[VARIANT]) -> Result<VARIANT> {
        log::info!("New window detected!");
        if params[0].Anonymous.Anonymous.vt != VT_BSTR.0 as u16 {
            return Err(E_FAIL.into());
        }

        let url = params[0].Anonymous.Anonymous.Anonymous.bstrVal.to_string();
        log::info!("New window url: {:?}", url);
        //self.tab_bar.new_window(url)?;
        Ok(Default::default())
    }

    #[allow(clippy::too_many_arguments)]
    pub unsafe fn Invoke(
        &mut self,
        dispidmember: i32,
        riid: *const Guid,
        lcid: u32,
        wflags: u16,
        pdispparams: *const DISPPARAMS,
        pvarresult: *mut VARIANT,
        pexcepinfo: *mut EXCEPINFO,
        puargerr: *mut u32,
    ) -> Result<()> {
        if wflags != DISPATCH_METHOD as u16 {
            return Err(DISP_E_MEMBERNOTFOUND.into());
        }

        let params = pdispparams.as_ref();
        let params = match params {
            Some(params) => std::slice::from_raw_parts(params.rgvarg, params.cArgs as usize),
            None => &[],
        };

        let result = match dispidmember {
            0xfc => self.NavigateComplete(params),
            0xfb => self.NewWindow(params),
            _ => Ok(Default::default()),
        };

        if let Ok(result) = result {
            if !pvarresult.is_null() {
                *pvarresult = result;
            }
        }

        Ok(())
    }
}

struct DeskBandData {
    //p_site: Rc<IUnknown>,
    p_input_object_site: Rc<IInputObjectSite>,

    tab_bar: Rc<tabs::tab_bar::TabBar>,
}

#[implement(Windows::Win32::System::Com::{IObjectWithSite},  Windows::Win32::UI::Shell::{IDeskBand})]
#[derive(Default)]
struct DeskBand {
    data: Option<DeskBandData>,
}

#[allow(non_snake_case)]
impl DeskBand {
    // IObjectWithSite
    pub unsafe fn SetSite(&mut self, unknown_site: &Option<IUnknown>) -> Result<()> {
        log::info!("Set Site");
        self.data = None;

        log::info!("Getting object site");
        let input_object_site: IInputObjectSite = unknown_site.as_ref().ok_or(E_FAIL)?.cast()?;

        log::info!("Acquiring services");
        let service_provider: IServiceProvider = input_object_site.cast()?;
        let web_browser =
            query_service_provider::<IWebBrowserApp>(&service_provider)?.cast::<IWebBrowser2>()?;
        let shell_browser = query_service_provider::<IShellBrowser>(&service_provider)?;

        log::info!("Creating tab bar");
        let parent_window_handle = unknown_site
            .as_ref()
            .ok_or(E_FAIL)?
            .cast::<IOleWindow>()?
            .GetWindow()?;

        let explorer_handle = shell_browser.GetWindow()?;

        let tab_bar = tabs::tab_bar::TabBar::new(
            parent_window_handle,
            explorer_handle,
            shell_browser.clone(),
        );
        tab_bar.add_tab(get_current_folder_pidl(&shell_browser).ok(), 0)?;

        log::info!("Connecting to event handler");

        let browser_event_handler = BrowserEventHandler {
            tab_bar: Rc::downgrade(&tab_bar),
            browser: shell_browser.clone(),
        };
        let container = web_browser.cast::<IConnectionPointContainer>()?;

        let iid = DWebBrowserEvents2::IID;
        let point = container.FindConnectionPoint(std::ptr::addr_of!(iid))?;

        point
            .Advise(IUnknown::from(browser_event_handler))
            .expect("advise failed");

        self.data = Some(DeskBandData {
            tab_bar,
            p_input_object_site: Rc::new(input_object_site),
        });

        detour::hook_browse_object(shell_browser);
        detour::hook_show_window(explorer_handle);

        log::info!("Set Site Ok");
        Ok(())
    }

    pub unsafe fn GetSite(&self, iid: *const Guid, out: *mut RawPtr) -> HRESULT {
        log::info!("Get site");
        if let Some(data) = &self.data {
            return data.p_input_object_site.query(iid, out);
        }
        E_FAIL
    }

    // IDeskBand
    pub unsafe fn GetWindow(&self) -> Result<HWND> {
        log::info!("Get window");
        if let Some(data) = &self.data {
            return Ok(data.tab_bar.get_handle());
        }
        Err(E_FAIL.into())
    }

    pub unsafe fn ContextSensitiveHelp(&self, _: BOOL) -> Result<()> {
        Err(E_NOTIMPL.into())
    }

    pub unsafe fn ShowDW(&self, _show: BOOL) -> Result<()> {
        log::info!("ShowDW {:?}", _show);
        if let Some(data) = &self.data {
            ShowWindow(
                data.tab_bar.get_handle(),
                match _show.0 {
                    0 => SW_HIDE,
                    _ => SW_SHOW,
                },
            );
        }
        Ok(())
    }

    pub unsafe fn CloseDW(&mut self, _reserved: u32) -> Result<()> {
        log::info!("CloseDW");
        if let Some(data) = &self.data {
            let handle = data.tab_bar.get_handle();
            ShowWindow(handle, SW_HIDE);
            DestroyWindow(handle);
        }
        self.data = None;
        Ok(())
    }

    pub unsafe fn ResizeBorderDW(
        &self,
        _prc_border: *const RECT,
        _unknown_toolbar_site: &Option<IUnknown>,
        _reserved: BOOL,
    ) -> Result<()> {
        log::info!("ResizeBorderDW");
        Err(E_NOTIMPL.into())
    }

    pub unsafe fn GetBandInfo(
        &self,
        _band_id: u32,
        _view_mode: u32,
        desk_band_info_ptr: *mut DESKBANDINFO,
    ) -> Result<()> {
        log::info!("GetBandInfo");
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
            desk_band_info.dwModeFlags = DBIMF_NORMAL | DBIMF_BKCOLOR | DBIMF_NOMARGINS;
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
            return CLASS_E_NOAGGREGATION;
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
        if _flock.as_bool() {
            DLL_LOCK += 1;
        } else {
            DLL_LOCK -= 1;
        }
        Ok(())
    }
}

// Dll stuff
#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllMain(instance: HINSTANCE, dw_reason: u32, _lpv_reserved: RawPtr) -> BOOL {
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
        unsafe { DisableThreadLibraryCalls(instance) };

        // Make this safe at some point
        unsafe {
            DLL_INSTANCE = Some(instance);
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
pub extern "stdcall" fn DllRegisterServer() -> HRESULT {
    S_OK
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "stdcall" fn DllUnregisterServer() -> HRESULT {
    S_OK
}
