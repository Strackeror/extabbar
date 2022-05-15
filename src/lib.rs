#![allow(clippy::forget_copy)]

mod detour;
mod idl;
mod settings;
mod tabs;

use std::ffi::c_void;
use std::path::PathBuf;
use std::rc::{Rc, Weak};
use std::sync::Mutex;

use once_cell::sync::Lazy;
use tabs::tab_bar::get_current_folder_path;
use windows::core::{implement, Result, GUID, PCWSTR};
use windows::Win32::System::LibraryLoader::DisableThreadLibraryCalls;
use windows::Win32::UI::Shell::{DWebBrowserEvents2, IWebBrowser2, IWebBrowserApp};
use windows::Win32::UI::WindowsAndMessaging::{
    DestroyWindow, EnumChildWindows, FindWindowExW, GetClassNameW, ShowWindow, SW_HIDE, SW_SHOW,
};
use Windows::core::{Abi, IUnknown, Interface, RawPtr, HRESULT, PWSTR};
use Windows::Win32::System::Com::{
    IClassFactory_Impl, IDispatch_Impl, ITypeInfo, DISPPARAMS, EXCEPINFO,
};
use Windows::Win32::System::LibraryLoader::GetModuleFileNameW;
use Windows::Win32::System::Ole::{
    IObjectWithSite_Impl, IOleWindow_Impl, DISPATCH_METHOD, VT_BSTR,
};

use windows::Win32::Foundation::*;
use windows::Win32::System::Com::{IConnectionPointContainer, IServiceProvider, VARIANT};
use windows::Win32::System::Ole::IOleWindow;
use windows::Win32::UI::Shell::*;

use crate::settings::current_settings;
use crate::tabs::tab_bar::DLL_INSTANCE;
use windows as Windows;

// {9ecce421-925a-4484-b2cf-c00b182bc32a}
pub const EXT_TAB_GUID: GUID = GUID::from_values(
    0x9ecce421,
    0x925a,
    0x4484,
    [0xb2, 0xcf, 0xc0, 0x0b, 0x18, 0x2b, 0xc3, 0x2a],
);

static mut DLL_LOCK: i32 = 0;
static MAIN_BAR_OPEN: Lazy<Mutex<bool>> = Lazy::new(|| Mutex::new(false));

pub const BROWSE_OBJECT_MESSAGE: &str = "extabbar_BrowseObject";
pub const SHOW_WINDOW_MESSAGE: &str = "extabbar_ShowWindow";

#[derive(Clone)]
struct BrowserEventHandlerContent {
    tab_bar: Weak<tabs::tab_bar::TabBar>,
    browser: IShellBrowser,
}
#[implement(DWebBrowserEvents2)]
struct BrowserEventHandler(std::sync::Mutex<BrowserEventHandlerContent>);

fn get_dll_path() -> PathBuf {
    unsafe {
        let mut output = [0u16; 256];
        let sz = GetModuleFileNameW(DLL_INSTANCE.unwrap(), &mut output) as usize;
        PathBuf::from(String::from_utf16_lossy(&output[..sz]))
    }
}

fn find_travel_toolbar(explorer_handle: HWND) -> Result<HWND> {
    let mut enum_output = HWND(0);
    unsafe extern "system" fn enum_proc(hwnd: HWND, param: LPARAM) -> BOOL {
        let mut wstr = [0u16; 256];
        let size = GetClassNameW(hwnd, &mut wstr) as usize;
        let string = String::from_utf16_lossy(&wstr[0..size]);

        // let string = match string {
        //     Err(_) => return BOOL(1),
        //     Ok(s) => s,
        // };
        log::info!("found class:{}", string);

        if string != "TravelBand" {
            return BOOL(1);
        }
        *(param.0 as *mut HWND) = hwnd;
        BOOL(0)
    }

    unsafe {
        EnumChildWindows(
            explorer_handle,
            Some(enum_proc),
            LPARAM(&mut enum_output as *mut _ as isize),
        );
        match FindWindowExW(enum_output, HWND(0), "ToolbarWindow32", PCWSTR::default()) {
            HWND(0) => Err(E_FAIL.into()),
            hwnd => Ok(hwnd),
        }
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

impl IDispatch_Impl for BrowserEventHandler {
    #[allow(clippy::too_many_arguments)]
    fn Invoke(
        &self,
        dispidmember: i32,
        _riid: *const GUID,
        _lcid: u32,
        wflags: u16,
        pdispparams: *const DISPPARAMS,
        pvarresult: *mut VARIANT,
        _pexcepinfo: *mut EXCEPINFO,
        _puargerr: *mut u32,
    ) -> Result<()> {
        if wflags != DISPATCH_METHOD as u16 {
            return Err(DISP_E_MEMBERNOTFOUND.into());
        }

        let params = unsafe { pdispparams.as_ref() };
        let params = match params {
            Some(params) => unsafe {
                std::slice::from_raw_parts(params.rgvarg, params.cArgs as usize)
            },
            None => &[],
        };
        let mut content = self.0.lock().unwrap();
        let result = match dispidmember {
            0xfc => content.navigate_complete(params),
            0xfb => content.new_window(params),
            _ => Ok(Default::default()),
        };

        if let Ok(result) = result {
            if !pvarresult.is_null() {
                unsafe {
                    *pvarresult = result;
                }
            }
        }

        Ok(())
    }

    fn GetTypeInfoCount(&self) -> Result<u32> {
        Ok(0)
    }

    fn GetTypeInfo(&self, _itinfo: u32, _lcid: u32) -> Result<ITypeInfo> {
        Err(E_NOTIMPL.into())
    }

    fn GetIDsOfNames(
        &self,
        _riid: *const GUID,
        rgsznames: *const PWSTR,
        _cnames: u32,
        lcid: u32,
        _rgdispid: *mut i32,
    ) -> Result<()> {
        log::info!("rgsznames:{:?} lcid:{:?}", unsafe { *rgsznames }, lcid);
        Err(DISP_E_UNKNOWNNAME.into())
    }
}

impl DWebBrowserEvents2_Impl for BrowserEventHandler {}

impl BrowserEventHandlerContent {
    fn navigate_complete(&self, _params: &[VARIANT]) -> Result<VARIANT> {
        let path = get_current_folder_path(&self.browser);
        self.tab_bar.upgrade().unwrap().navigated(path)?;

        Ok(Default::default())
    }

    fn new_window(&mut self, params: &[VARIANT]) -> Result<VARIANT> {
        log::info!("New window detected!");
        if unsafe { params[0].Anonymous.Anonymous.vt } != VT_BSTR.0 as u16 {
            return Err(E_FAIL.into());
        }

        let url = unsafe { params[0].Anonymous.Anonymous.Anonymous.bstrVal.to_string() };
        log::info!("New window url: {:?}", url);
        //self.tab_bar.new_window(url)?;
        Ok(Default::default())
    }
}

struct DeskBandData {
    //p_site: Rc<IUnknown>,
    p_input_object_site: Rc<IInputObjectSite>,

    tab_bar: Rc<tabs::tab_bar::TabBar>,
}

#[implement(
    Windows::Win32::System::Ole::IObjectWithSite,
    Windows::Win32::UI::Shell::IDeskBand
)]
#[derive(Default)]
struct DeskBand {
    data: Mutex<Option<DeskBandData>>,
}
impl IOleWindow_Impl for DeskBand {
    fn GetWindow(&self) -> Result<HWND> {
        log::info!("Get window");
        if let Some(data) = &*self.data.lock().unwrap() {
            return Ok(data.tab_bar.get_handle());
        }
        Err(E_FAIL.into())
    }

    fn ContextSensitiveHelp(&self, _: BOOL) -> Result<()> {
        Err(E_NOTIMPL.into())
    }
}

impl IDockingWindow_Impl for DeskBand {
    fn ShowDW(&self, _show: BOOL) -> Result<()> {
        log::info!("ShowDW {:?}", _show);
        if let Some(data) = &*self.data.lock().unwrap() {
            unsafe {
                ShowWindow(
                    data.tab_bar.get_handle(),
                    match _show.0 {
                        0 => SW_HIDE,
                        _ => SW_SHOW,
                    },
                );
            }
        }
        Ok(())
    }

    fn CloseDW(&self, _reserved: u32) -> Result<()> {
        log::info!("CloseDW");
        if let Some(data) = &*self.data.lock().unwrap() {
            let handle = data.tab_bar.get_handle();
            if data.tab_bar.is_main() {
                *MAIN_BAR_OPEN.lock().unwrap() = false;
            }
            unsafe {
                ShowWindow(handle, SW_HIDE);
                DestroyWindow(handle);
            }
        }
        *self.data.lock().unwrap() = None;
        Ok(())
    }

    fn ResizeBorderDW(
        &self,
        _prc_border: *const RECT,
        _unknown_toolbar_site: &Option<IUnknown>,
        _reserved: BOOL,
    ) -> Result<()> {
        log::info!("ResizeBorderDW");
        Err(E_NOTIMPL.into())
    }
}

impl IDeskBand_Impl for DeskBand {
    fn GetBandInfo(
        &self,
        _band_id: u32,
        _view_mode: u32,
        desk_band_info_ptr: *mut DESKBANDINFO,
    ) -> Result<()> {
        log::info!("GetBandInfo");
        if desk_band_info_ptr.is_null() {
            return E_INVALIDARG.ok();
        }

        log::info!("get band info id:{}, view mode:{}", _band_id, _view_mode);

        let desk_band_info = unsafe { desk_band_info_ptr.as_mut() }.ok_or(E_INVALIDARG)?;
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

impl IObjectWithSite_Impl for DeskBand {
    fn SetSite(&self, unknown_site: &Option<IUnknown>) -> Result<()> {
        log::info!(
            "Set Site, data active:{:?}",
            self.data.lock().unwrap().is_some()
        );
        *self.data.lock().unwrap() = None;

        log::info!("Getting object site");
        let input_object_site: IInputObjectSite = unknown_site.as_ref().ok_or(E_FAIL)?.cast()?;

        log::info!("Acquiring services");
        let service_provider: IServiceProvider = input_object_site.cast()?;
        let web_browser =
            query_service_provider::<IWebBrowserApp>(&service_provider)?.cast::<IWebBrowser2>()?;
        let shell_browser = query_service_provider::<IShellBrowser>(&service_provider)?;

        log::info!("Creating tab bar");
        let parent_window_handle = unsafe {
            unknown_site
                .as_ref()
                .ok_or(E_FAIL)?
                .cast::<IOleWindow>()?
                .GetWindow()?
        };

        let browser_handle = unsafe { HWND(web_browser.HWND()?.0 as _) };
        let travel_toolbar_handle = find_travel_toolbar(browser_handle)?;

        let explorer_handle = unsafe { shell_browser.GetWindow()? };

        let settings = current_settings();
        let mut is_main = false;
        {
            let mut bar_open = MAIN_BAR_OPEN.lock().unwrap();
            if !*bar_open {
                *bar_open = true;
                is_main = true;
            }
        }
        let tab_bar = tabs::tab_bar::TabBar::new(
            parent_window_handle,
            explorer_handle,
            travel_toolbar_handle,
            shell_browser.clone(),
            settings,
            is_main,
        );

        tab_bar.add_tab(get_current_folder_path(&shell_browser), 0)?;

        log::info!("Connecting to event handler");
        let browser_event_handler = BrowserEventHandler(Mutex::new(BrowserEventHandlerContent {
            tab_bar: Rc::downgrade(&tab_bar),
            browser: shell_browser.clone(),
        }));
        let container = web_browser.cast::<IConnectionPointContainer>()?;

        let iid = DWebBrowserEvents2::IID;
        let point = unsafe { container.FindConnectionPoint(std::ptr::addr_of!(iid))? };

        unsafe {
            point
                .Advise(IUnknown::from(browser_event_handler))
                .expect("advise failed");
        }

        *self.data.lock().unwrap() = Some(DeskBandData {
            tab_bar,
            p_input_object_site: Rc::new(input_object_site),
        });

        unsafe {
            detour::hook_browse_object(shell_browser);
            detour::hook_show_window();
            if is_main {
                detour::set_main_explorer(explorer_handle);
            }
        }

        log::info!("Set Site Ok");
        Ok(())
    }

    fn GetSite(&self, iid: *const GUID, out: *mut RawPtr) -> Result<()> {
        log::info!("Get site");

        match &*self.data.lock().unwrap() {
            Some(data) => unsafe { data.p_input_object_site.query(&*iid, out) },
            None => E_FAIL,
        }
        .ok()
    }
}

#[allow(non_snake_case)]
impl DeskBand {}

#[implement(Windows::Win32::System::Com::IClassFactory)]
struct ClassFactory {}

#[allow(non_snake_case)]
impl IClassFactory_Impl for ClassFactory {
    fn CreateInstance(
        &self,
        outer: &Option<IUnknown>,
        iid: *const GUID,
        object: *mut RawPtr,
    ) -> Result<()> {
        if outer.is_some() {
            return CLASS_E_NOAGGREGATION.ok();
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
            deskband_unknown.query(&*iid, object).ok()
        }
    }

    fn LockServer(&self, _flock: BOOL) -> Result<()> {
        unsafe {
            if _flock.as_bool() {
                DLL_LOCK += 1;
            } else {
                DLL_LOCK -= 1;
            }
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

        // Make this safe at some point
        unsafe {
            DLL_INSTANCE = Some(instance);
        }

        let current_path = get_dll_path().with_file_name("extabbar.log");
        fern::Dispatch::new()
            .level(log::LevelFilter::Debug)
            .format(|out, message, record| {
                out.finish(format_args!(
                    "{}[{:40}][{:5}] {}",
                    chrono::Local::now().format("[%Y-%m-%d-%H:%M:%S]"),
                    record.target(),
                    record.level(),
                    message
                ))
            })
            .chain(fern::log_file(current_path.clone()).unwrap())
            .apply()
            .unwrap();
        log::info!("Attached, dll path: {:?}", current_path);
        std::panic::set_hook(Box::new(|info| log::error!("PANIC ! {:?}", info)));
        unsafe { DisableThreadLibraryCalls(instance) };
    }
    true.into()
}

/// # Safety
#[no_mangle]
#[allow(non_snake_case)]
pub unsafe extern "stdcall" fn DllGetClassObject(
    rclsid: *const GUID,
    iid: *const GUID,
    object: *mut RawPtr,
) -> HRESULT {
    if EXT_TAB_GUID == *rclsid {
        log::info!("Dll Got ClassObject");
        let unknown: IUnknown = ClassFactory {}.into();
        return unknown.query(&*iid, object);
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
