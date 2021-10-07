fn main() {
    windows::build! {
      Windows::Win32::Foundation::*,
      Windows::Win32::System::Com::{IClassFactory, IPersistStream, IObjectWithSite, IConnectionPoint, IConnectionPointContainer},
      Windows::Win32::System::LibraryLoader::DisableThreadLibraryCalls,
      Windows::Win32::System::OleAutomation::IDispatch,
      Windows::Win32::System::SystemServices::IServiceProvider,
      Windows::Win32::System::WindowsProgramming::{IWebBrowser2, IWebBrowserApp, DWebBrowserEvents2},
      Windows::Win32::UI::Shell::*,
      Windows::Win32::UI::Controls::*,
      Windows::Win32::UI::WindowsAndMessaging::*,
    }
}
