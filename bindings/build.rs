fn main() {
    windows::build! {
      Windows::Win32::Foundation::*,
      Windows::Win32::System::Com::{IClassFactory, IPersistStream, IObjectWithSite},
      Windows::Win32::UI::Shell::*,
      Windows::Win32::UI::Controls::*,
      Windows::Win32::UI::WindowsAndMessaging::*,
    }
}
