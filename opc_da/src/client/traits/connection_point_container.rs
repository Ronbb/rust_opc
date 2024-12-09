use windows::core::GUID;
use windows::Win32::System::Com::IConnectionPoint;

pub trait ConnectionPointContainerTrait {
    fn interface(
        &self,
    ) -> windows::core::Result<&windows::Win32::System::Com::IConnectionPointContainer>;

    fn find_connection_point(&self, id: &GUID) -> windows::core::Result<IConnectionPoint> {
        unsafe { self.interface()?.FindConnectionPoint(id) }
    }

    fn enum_connection_points(
        &self,
    ) -> windows::core::Result<windows::Win32::System::Com::IEnumConnectionPoints> {
        unsafe { self.interface()?.EnumConnectionPoints() }
    }
}
